# -*- coding: utf-8 -*-
"""
承包方家庭成员 & 承包地块信息提取工具
从 D:\农经全延保 下所有表1/表2 xlsx 文件中提取信息，汇总到一张表。
表1：同一户中同一人如在确权区和变动区都出现，合并为一条记录。减少的人口自动删除。
表2：提取地块确权登记信息及地块增减变动情况。
"""

import os
import re
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment


# ── 常量 ──────────────────────────────────────────────────────────────────────

OUTPUT_COLUMNS = (
    "所属组", "编号", "承包方编码", "承包方代表",
    "发包方名称", "户内序号",
    "家庭成员姓名", "性别", "身份证号",
    "与承包方代表关系", "变动情况",
)

OUTPUT_COLUMNS_B2 = (
    "所属组", "编号", "承包方编码", "承包方代表",
    "联系方式", "确权总面积(亩)",
    "地块总数", "地块序号",
    "地块名称", "地块编码", "地块面积(亩)",
    "东至", "西至", "南至", "北至",
    "变动情况", "变动面积(亩)", "变动原因",
)

_HOUSEHOLD_KEY_COLS = (1, 2)

_REDUCE_REMOVE_KEYWORDS = ("嫁", "死", "亡", "去世", "登记错误")

_FILL_A = PatternFill(start_color="DAEEF3", end_color="DAEEF3", fill_type="solid")
_FILL_B = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
_SEP_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")


# ── 工具函数 ──────────────────────────────────────────────────────────────────

def _parse_filename(filename):
    # 先匹配 编号-姓名-表1.xlsx
    m = re.match(r"^(.+?)-(.+?)-表1\.xlsx$", filename)
    if m:
        return m.group(1), m.group(2)
    # 再匹配 编号-姓名表1.xlsx（只有一个横线）
    m = re.match(r"^(.+?)-(.+?)表1\.xlsx$", filename)
    if m:
        return m.group(1), m.group(2)
    # 无横线时整体作为编号
    base = filename.replace("表1.xlsx", "")
    return base.rstrip("-"), ""


def _parse_filename2(filename):
    # 先匹配 编号-姓名-表2.xlsx
    m = re.match(r"^(.+?)-(.+?)-表2\.xlsx$", filename)
    if m:
        return m.group(1), m.group(2)
    # 再匹配 编号-姓名表2.xlsx（只有一个横线）
    m = re.match(r"^(.+?)-(.+?)表2\.xlsx$", filename)
    if m:
        return m.group(1), m.group(2)
    # 无横线时整体作为编号
    base = filename.replace("表2.xlsx", "")
    return base.rstrip("-"), ""


def _normalize(text):
    return str(text).replace(" ", "").replace("\n", "").replace("\r", "")


def _find_section_header(ws, keyword):
    for r in range(1, ws.max_row + 1):
        for c in range(1, 6):
            val = ws.cell(r, c).value
            if val and keyword in _normalize(val):
                return r
    return None


def _is_seq_number(val):
    if val is None:
        return False
    if isinstance(val, (int, float)):
        return val > 0
    s = str(val).strip()
    try:
        return float(s) > 0
    except (ValueError, TypeError):
        return False


# ── 表1 数据读取 ──────────────────────────────────────────────────────────────

def _read_section_data(ws, header_row, end_row=None, include_change_col=False):
    results = []
    max_r = (end_row - 1) if end_row else ws.max_row
    for r in range(header_row + 1, max_r + 1):
        seq = ws.cell(r, 2).value
        name = ws.cell(r, 4).value
        if not name or str(name).strip() == "":
            continue
        if not _is_seq_number(seq):
            continue
        row_data = {
            "家庭成员姓名": str(name).strip(),
            "性别": str(ws.cell(r, 6).value or "").strip(),
            "身份证号": str(ws.cell(r, 7).value or "").strip(),
            "与承包方代表关系": str(ws.cell(r, 9).value or "").strip(),
            "变动情况": str(ws.cell(r, 3).value or "").strip() if include_change_col else "",
        }
        if include_change_col:
            row_data["_reason"] = str(ws.cell(r, 13).value or "").strip()
        results.append(row_data)
    return results


def _person_key(row):
    id_num = row.get("身份证号", "")
    if id_num and id_num != "/":
        return ("id", id_num)
    return ("name", row.get("家庭成员姓名", ""))


def _should_remove(row):
    if row.get("变动情况", "") != "减少":
        return False
    return True


# ── 表1 解析 ──────────────────────────────────────────────────────────────────

def parse_biao1(filepath, group_name):
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=False)
    ws = wb.active

    fn = os.path.basename(filepath)
    file_code, file_name = _parse_filename(fn)

    contractor = str(ws.cell(4, 5).value or file_name or "").strip()
    fa_bao_fang = str(ws.cell(3, 1).value or "").strip()
    fa_bao_fang = fa_bao_fang.replace("发包方名称：", "")
    fa_bao_fang = fa_bao_fang.replace("发包方名称:", "").strip()

    # 提取承包方编码：确权承包合同编号去掉末尾字母
    contract_no = str(ws.cell(5, 9).value or "").strip()
    contractor_code = contract_no.rstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz") if contract_no else (file_code or "")

    info = {
        "所属组": group_name,
        "编号": file_code or "",
        "承包方编码": contractor_code,
        "承包方代表": contractor,
        "发包方名称": fa_bao_fang,
    }

    h1 = _find_section_header(ws, "确权")
    h2 = _find_section_header(ws, "变动")

    section1 = _read_section_data(ws, h1, end_row=h2, include_change_col=False) if h1 else []
    section2 = _read_section_data(ws, h2, include_change_col=True) if h2 else []
    wb.close()

    s1_key_idx = {}
    for i, row in enumerate(section1):
        s1_key_idx[_person_key(row)] = i

    s2_only = []
    for row in section2:
        pk = _person_key(row)
        if pk in s1_key_idx:
            idx = s1_key_idx[pk]
            section1[idx]["变动情况"] = row["变动情况"]
            section1[idx]["_reason"] = row.get("_reason", "")
            # 变动区出现"本人"属户主变更，关系按变更后的值填写
            new_rel = row.get("与承包方代表关系", "").strip()
            if new_rel:
                section1[idx]["与承包方代表关系"] = new_rel
                if new_rel == "本人":
                    info["承包方代表"] = row["家庭成员姓名"]
        else:
            s2_only.append(row)

    merged = section1 + s2_only

    filtered = []
    for row in merged:
        if _should_remove(row):
            continue
        row.pop("_reason", None)
        filtered.append(row)
    return info, filtered


# ── 表2 解析 ──────────────────────────────────────────────────────────────────

def parse_biao2(filepath, group_name):
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=False)
    ws = wb.active

    fn = os.path.basename(filepath)
    file_code, file_name = _parse_filename2(fn)

    contractor = str(ws.cell(4, 4).value or file_name or "").strip()
    phone = str(ws.cell(4, 6).value or "").strip()
    total_area = str(ws.cell(4, 10).value or "").strip()
    total_plots = str(ws.cell(4, 14).value or "").strip()

    # 承包方编码：从文件名前缀提取
    contractor_code = file_code or ""

    info = {
        "所属组": group_name,
        "编号": file_code or "",
        "承包方编码": contractor_code,
        "承包方代表": contractor,
        "联系方式": phone,
        "确权总面积(亩)": total_area,
        "地块总数": total_plots,
    }

    h1 = _find_section_header(ws, "确权")
    h2 = _find_section_header(ws, "变动")

    # Read confirmed plots
    plots = []
    if h1:
        end = (h2 - 1) if h2 else ws.max_row
        for r in range(h1 + 1, end + 1):
            name = ws.cell(r, 3).value
            if not name or str(name).strip() == "":
                continue
            plots.append({
                "地块名称": str(name).strip(),
                "地块编码": str(ws.cell(r, 4).value or "").strip(),
                "地块面积(亩)": str(ws.cell(r, 6).value or "").strip(),
                "东至": str(ws.cell(r, 7).value or "").strip(),
                "西至": str(ws.cell(r, 8).value or "").strip(),
                "南至": str(ws.cell(r, 9).value or "").strip(),
                "北至": str(ws.cell(r, 10).value or "").strip(),
            })

    # Read changes
    changes = {}
    if h2:
        for r in range(h2 + 1, ws.max_row + 1):
            name = ws.cell(r, 4).value
            if not name or str(name).strip() == "":
                continue
            changes[str(name).strip()] = {
                "变动情况": str(ws.cell(r, 3).value or "").strip(),
                "变动面积(亩)": str(ws.cell(r, 11).value or "").strip(),
                "变动原因": str(ws.cell(r, 13).value or "").strip(),
            }
    wb.close()

    # Merge
    rows = []
    for i, p in enumerate(plots, 1):
        ch = changes.pop(p["地块名称"], None)
        row = {
            "地块序号": i,
            "地块名称": p["地块名称"],
            "地块编码": p["地块编码"],
            "地块面积(亩)": p["地块面积(亩)"],
            "东至": p["东至"], "西至": p["西至"],
            "南至": p["南至"], "北至": p["北至"],
        }
        if ch:
            row["变动情况"] = ch["变动情况"]
            row["变动面积(亩)"] = ch["变动面积(亩)"]
            row["变动原因"] = ch["变动原因"]
        else:
            row["变动情况"] = "无"
            row["变动面积(亩)"] = ""
            row["变动原因"] = ""
        rows.append(row)

    # Newly added plots from changes
    for ch in changes.values():
        rows.append({
            "地块序号": len(rows) + 1,
            "地块名称": "",
            "地块编码": "",
            "地块面积(亩)": "",
            "东至": "", "西至": "",
            "南至": "", "北至": "",
            "变动情况": ch["变动情况"],
            "变动面积(亩)": ch["变动面积(亩)"],
            "变动原因": ch["变动原因"],
        })

    return info, rows


def _household_key(row):
    return (row[_HOUSEHOLD_KEY_COLS[0]], row[_HOUSEHOLD_KEY_COLS[1]])


# ── 扫描 ──────────────────────────────────────────────────────────────────────

def scan_folder(folder_path):
    results = []
    errors = []
    xlsx_files = []
    for root, dirs, files in os.walk(folder_path):
        for f in files:
            if f.endswith("表1.xlsx") and not f.startswith("~$"):
                rel = os.path.relpath(root, folder_path)
                group = rel if rel != "." else os.path.basename(folder_path)
                xlsx_files.append((os.path.join(root, f), group))
    xlsx_files.sort(key=lambda x: x[0])
    total = len(xlsx_files)

    for idx, (fp, group) in enumerate(xlsx_files):
        try:
            info, rows = parse_biao1(fp, group)
            for seq, row in enumerate(rows, 1):
                merged = {**info, **row}
                merged["户内序号"] = seq
                c = tuple(merged.get(col, "") for col in OUTPUT_COLUMNS)
                results.append({"values": c, "key": _household_key(c)})
        except Exception as e:
            errors.append("%s: %s" % (os.path.basename(fp), str(e)))
        yield idx + 1, total, results, errors


def scan_folder_b2(folder_path):
    results = []
    errors = []
    xlsx_files = []
    for root, dirs, files in os.walk(folder_path):
        for f in files:
            if f.endswith("表2.xlsx") and not f.startswith("~$"):
                rel = os.path.relpath(root, folder_path)
                group = rel if rel != "." else os.path.basename(folder_path)
                xlsx_files.append((os.path.join(root, f), group))
    xlsx_files.sort(key=lambda x: x[0])
    total = len(xlsx_files)

    for idx, (fp, group) in enumerate(xlsx_files):
        try:
            info, rows = parse_biao2(fp, group)
            for row in rows:
                merged = {**info, **row}
                c = tuple(merged.get(col, "") for col in OUTPUT_COLUMNS_B2)
                results.append({"values": c, "key": _household_key(c)})
        except Exception as e:
            errors.append("%s: %s" % (os.path.basename(fp), str(e)))
        yield idx + 1, total, results, errors


# ── GUI ───────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    _tag_colors = {"h0": "#DAEEF3", "h1": "#E2EFDA", "sep": "#F2F2F2"}

    def __init__(self):
        super().__init__()
        self.title("承包方家庭成员 & 承包地块提取工具")
        self.geometry("1200x650")
        self.minsize(900, 500)
        self.results_b1 = []
        self.results_b2 = []
        self._scanning = False
        self._build_ui()

    def _build_ui(self):
        style = ttk.Style(self)
        style.configure("TButton", padding=(10, 4))
        style.configure("TLabel", padding=(10, 4))
        style.configure("Big.TButton", padding=(10, 6))

        # 文件夹
        top = ttk.Frame(self, padding=(10, 8, 10, 4))
        top.pack(fill=tk.X)
        ttk.Label(top, text="源文件夹：").pack(side=tk.LEFT)
        self.folder_var = tk.StringVar(value=r"D:\农经全延保")
        ttk.Entry(top, textvariable=self.folder_var, width=60).pack(side=tk.LEFT, padx=(0, 6), fill=tk.X, expand=True)
        ttk.Button(top, text="浏览…", command=self._browse).pack(side=tk.LEFT)

        # 按钮行
        btn_frame = ttk.Frame(self, padding=(10, 4))
        btn_frame.pack(fill=tk.X)
        self.start_btn = ttk.Button(btn_frame, text="提取", style="Big.TButton", command=self._start)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 12))
        self.export_btn = ttk.Button(btn_frame, text="导出 Excel", command=self._export, state=tk.DISABLED)
        self.export_btn.pack(side=tk.LEFT, padx=(0, 12))
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(btn_frame, textvariable=self.status_var).pack(side=tk.LEFT, padx=(12, 0))

        # 进度条
        prog_frame = ttk.Frame(self, padding=(10, 2, 10, 4))
        prog_frame.pack(fill=tk.X)
        self.progress = ttk.Progressbar(prog_frame, mode="determinate", length=400)
        self.progress.pack(fill=tk.X, expand=True)

        # Notebook
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(4, 0))

        # Tab 1
        tab1 = ttk.Frame(self.notebook)
        self.notebook.add(tab1, text="  家庭成员  ")
        self.tree_b1 = ttk.Treeview(tab1, columns=list(range(len(OUTPUT_COLUMNS))), show="headings", height=16)
        self._setup_tree(self.tree_b1, OUTPUT_COLUMNS)
        vsb1 = ttk.Scrollbar(tab1, orient=tk.VERTICAL, command=self.tree_b1.yview)
        hsb1 = ttk.Scrollbar(tab1, orient=tk.HORIZONTAL, command=self.tree_b1.xview)
        self.tree_b1.configure(yscrollcommand=vsb1.set, xscrollcommand=hsb1.set)
        self.tree_b1.grid(row=0, column=0, sticky="nsew")
        vsb1.grid(row=0, column=1, sticky="ns")
        hsb1.grid(row=1, column=0, sticky="ew")
        tab1.rowconfigure(0, weight=1)
        tab1.columnconfigure(0, weight=1)

        # Tab 2
        tab2 = ttk.Frame(self.notebook)
        self.notebook.add(tab2, text="  承包地块  ")
        self.tree_b2 = ttk.Treeview(tab2, columns=list(range(len(OUTPUT_COLUMNS_B2))), show="headings", height=16)
        self._setup_tree(self.tree_b2, OUTPUT_COLUMNS_B2)
        vsb2 = ttk.Scrollbar(tab2, orient=tk.VERTICAL, command=self.tree_b2.yview)
        hsb2 = ttk.Scrollbar(tab2, orient=tk.HORIZONTAL, command=self.tree_b2.xview)
        self.tree_b2.configure(yscrollcommand=vsb2.set, xscrollcommand=hsb2.set)
        self.tree_b2.grid(row=0, column=0, sticky="nsew")
        vsb2.grid(row=0, column=1, sticky="ns")
        hsb2.grid(row=1, column=0, sticky="ew")
        tab2.rowconfigure(0, weight=1)
        tab2.columnconfigure(0, weight=1)

        # 底部
        bottom = ttk.Frame(self, padding=(10, 2, 10, 8))
        bottom.pack(fill=tk.X)
        self.count_var = tk.StringVar(value="")
        ttk.Label(bottom, textvariable=self.count_var, font=("", 10)).pack(side=tk.LEFT)

    def _setup_tree(self, tree, columns):
        col_widths = {
            "所属组": 60, "编号": 170, "承包方编码": 200, "承包方代表": 90,
            "发包方名称": 160, "户内序号": 50,
            "家庭成员姓名": 90, "性别": 50, "身份证号": 180,
            "与承包方代表关系": 120, "变动情况": 70,
            "联系方式": 110, "确权总面积(亩)": 90,
            "地块总数": 60, "地块序号": 60,
            "地块名称": 90, "地块编码": 180,
            "地块面积(亩)": 90,
            "东至": 150, "西至": 150, "南至": 150, "北至": 150,
            "变动面积(亩)": 90, "变动原因": 180,
        }
        for i, name in enumerate(columns):
            tree.heading(i, text=name)
            tree.column(i, width=col_widths.get(name, 80), minwidth=40)
        for tag, color in self._tag_colors.items():
            tree.tag_configure(tag, background=color)

    def _browse(self):
        d = filedialog.askdirectory(initialdir=self.folder_var.get())
        if d:
            self.folder_var.set(d)

    def _lock(self):
        self.start_btn.configure(state=tk.DISABLED)
        self.export_btn.configure(state=tk.DISABLED)

    def _unlock(self):
        self.start_btn.configure(state=tk.NORMAL)
        if self.results_b1 or self.results_b2:
            self.export_btn.configure(state=tk.NORMAL)

    # ── 统一提取 ──

    def _start(self):
        folder = self.folder_var.get().strip()
        if not os.path.isdir(folder):
            messagebox.showerror('错误', '请选择有效的文件夹。')
            return
        if self._scanning:
            return
        self._scanning = True
        self._lock()
        self.tree_b1.delete(*self.tree_b1.get_children())
        self.tree_b2.delete(*self.tree_b2.get_children())
        self.results_b1 = []
        self.results_b2 = []
        self._all_errors = []
        self.progress['value'] = 0
        self.status_var.set('正在提取…')
        threading.Thread(target=self._run, args=(folder,), daemon=True).start()

    def _run(self, folder):
        # 表1
        results = []
        errors = []
        done = 0
        total = 0
        for done, total, results, errors in scan_folder(folder):
            self.after(0, self._update_progress, done, total * 2, '家庭成员')
        self.results_b1 = results
        self._all_errors.extend(errors)
        b1_total = done if done else 1
        # 表2
        results = []
        errors = []
        done = 0
        for done, total, results, errors in scan_folder_b2(folder):
            self.after(0, self._update_progress, b1_total + done, b1_total + total, '承包地块')
        self.results_b2 = results
        self._all_errors.extend(errors)
        self.after(0, self._finish)

    def _finish(self):
        self._scanning = False
        self._fill_tree(self.tree_b1, self.results_b1, OUTPUT_COLUMNS)
        self._fill_tree(self.tree_b2, self.results_b2, OUTPUT_COLUMNS_B2, is_b2=True)
        self._unlock()
        hk1 = len(set(i['key'] for i in self.results_b1))
        hk2 = len(set(i['key'] for i in self.results_b2))
        self.status_var.set('提取完成')
        self.count_var.set(
            '家庭成员：%d 条(%d 户) | 承包地块：%d 条(%d 户)'
            % (len(self.results_b1), hk1, len(self.results_b2), hk2))
        msg = '提取完成！'
        msg += '\n家庭成员：%d 条(%d 户)' % (len(self.results_b1), hk1)
        msg += '\n承包地块：%d 条(%d 户)' % (len(self.results_b2), hk2)
        if self._all_errors:
            msg += '\n\n错误 %d 条' % len(self._all_errors)
        messagebox.showinfo('完成', msg)


    # ── 公用 ──

    def _update_progress(self, done, total, label):
        self.progress["maximum"] = total
        self.progress["value"] = done
        self.status_var.set("正在提取%s %d/%d …" % (label, done, total))

    # b2 household info columns that should only show once per group
    _B2_HOUSEHOLD_COLS = frozenset(range(7))  # cols 0-6: 所属组, 编号, 承包方编码, 承包方代表, 联系方式, 确权总面积, 地块总数

    def _fill_tree(self, tree, results, columns, is_b2=False):
        tree.delete(*tree.get_children())
        if not results:
            return
        prev_key = None
        parity = 0
        for item in results:
            cur_key = item["key"]
            if prev_key is not None and cur_key != prev_key:
                parity = 1 - parity
                sep = [""] * len(columns)
                sep[0] = "───"
                tree.insert("", tk.END, values=sep, tags=("sep",))
            tag = "h%d" % parity
            vals = list(item["values"])
            if is_b2 and prev_key == cur_key:
                for ci in self._B2_HOUSEHOLD_COLS:
                    vals[ci] = ""
            tree.insert("", tk.END, values=vals, tags=(tag,))
            prev_key = cur_key

    def _export(self):
        has_b1 = bool(self.results_b1)
        has_b2 = bool(self.results_b2)
        if not has_b1 and not has_b2:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="汇总表.xlsx",
            initialdir=self.folder_var.get(),
        )
        if not path:
            return
        try:
            wb = openpyxl.Workbook()
            if has_b1:
                ws1 = wb.active
                self._export_sheet(ws1, self.results_b1, list(OUTPUT_COLUMNS), "家庭成员")
            if has_b2:
                ws2 = wb.create_sheet() if has_b1 else wb.active
                self._export_sheet(ws2, self.results_b2, list(OUTPUT_COLUMNS_B2), "承包地块", is_b2=True)
            wb.save(path)
            messagebox.showinfo("导出成功", "已保存到：\n%s" % path)
        except Exception as e:
            messagebox.showerror("导出失败", str(e))

    def _export_sheet(self, ws, results, cols, title, is_b2=False):
        ws.title = title
        center_align = Alignment(horizontal="center", vertical="center")
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, size=11, color="FFFFFF")
        thick_bottom = Border(bottom=Side(style="medium", color="808080"))

        for ci, cn in enumerate(cols, 1):
            cell = ws.cell(1, ci, cn)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align

        cur_row = 2
        prev_key = None
        parity = 0
        for item in results:
            vals = list(item["values"])
            cur_key = item["key"]
            if is_b2 and prev_key == cur_key:
                for ci in range(7):
                    vals[ci] = ""
            if prev_key is not None and cur_key != prev_key:
                parity = 1 - parity
                for c in range(1, len(cols) + 1):
                    ws.cell(cur_row - 1, c).border = thick_bottom
                for c in range(1, len(cols) + 1):
                    ws.cell(cur_row, c, "").fill = _SEP_FILL
                ws.row_dimensions[cur_row].height = 6
                cur_row += 1
            fill = _FILL_A if parity == 0 else _FILL_B
            for ci, val in enumerate(vals, 1):
                cell = ws.cell(cur_row, ci, val)
                cell.fill = fill
                if ci <= 7:
                    cell.alignment = center_align
            prev_key = cur_key
            cur_row += 1

        for c in range(1, len(cols) + 1):
            ws.cell(cur_row - 1, c).border = thick_bottom

        for ci, cn in enumerate(cols, 1):
            max_len = len(cn)
            for r in range(2, min(cur_row, 500)):
                cv = str(ws.cell(r, ci).value or "")
                cl = sum(2 if ord(ch) > 127 else 1 for ch in cv)
                max_len = max(max_len, cl)
            ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = min(max_len + 3, 40)
        ws.freeze_panes = "A2"


if __name__ == "__main__":
    app = App()
    app.mainloop()