"""
批量文字替换工具 - 对文件夹内所有可编辑文件进行文字替换。
先复制源文件夹，在副本上操作，保留原文件格式。
支持: Word(.doc/.docx), Excel(.xls/.xlsx), PPT(.pptx), 纯文本文件
"""

import os
import shutil
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

# ── 可处理的文本文件扩展名 ──────────────────────────────────────────
TEXT_EXTENSIONS = {
    ".txt", ".csv", ".json", ".xml", ".html", ".htm", ".md", ".markdown",
    ".py", ".js", ".ts", ".jsx", ".tsx", ".css", ".scss", ".less",
    ".java", ".c", ".cpp", ".h", ".hpp", ".cs", ".go", ".rs", ".rb",
    ".php", ".swift", ".kt", ".scala", ".r", ".m", ".sh", ".bat",
    ".ps1", ".yaml", ".yml", ".toml", ".ini", ".cfg", ".conf",
    ".sql", ".log", ".tex", ".rst", ".svg",
}

SKIP_EXTENSIONS = {
    ".pdf", ".exe", ".dll", ".so", ".dylib",
    ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".ico", ".tiff", ".webp",
    ".mp3", ".mp4", ".avi", ".mkv", ".mov", ".wmv", ".flv", ".wav",
    ".zip", ".rar", ".7z", ".tar", ".gz", ".bz2",
    ".ppt",  # 旧格式 ppt 暂不支持
    ".pyc", ".class", ".o", ".obj", ".lib",
    ".ttf", ".otf", ".woff", ".woff2", ".eot",
    ".db", ".sqlite", ".sqlite3", ".mdb",
    ".bin", ".dat", ".pkl", ".pickle",
}


# ── 核心替换逻辑 ────────────────────────────────────────────────────

def copy_folder(src, dst):
    """复制源文件夹到目标路径，dst 必须不存在。"""
    shutil.copytree(src, dst)


def replace_in_text_file(filepath, find_text, replace_text):
    """对纯文本文件执行替换，返回替换次数。"""
    try:
        content = Path(filepath).read_text(encoding="utf-8")
    except (UnicodeDecodeError, UnicodeError):
        try:
            content = Path(filepath).read_text(encoding="gbk")
        except Exception:
            return 0

    if find_text not in content:
        return 0

    count = content.count(find_text)
    new_content = content.replace(find_text, replace_text)

    enc = "utf-8"
    try:
        content.encode("gbk")
        new_content.encode("gbk")
        enc = "gbk"
    except Exception:
        pass

    Path(filepath).write_text(new_content, encoding=enc)
    return count


def _replace_in_runs(runs, find_text, replace_text):
    """替换一段 runs 中的文字，处理中文跨 run 拆分的情况。返回替换次数。"""
    full_text = "".join(run.text or "" for run in runs)
    if find_text not in full_text:
        return 0

    occurrence_count = full_text.count(find_text)

    # 先尝试逐 run 替换（可保留格式）
    for run in runs:
        if find_text in (run.text or ""):
            run.text = run.text.replace(find_text, replace_text)

    # 检查是否全部替换成功
    remaining = "".join(run.text or "" for run in runs)
    if find_text not in remaining:
        return occurrence_count

    # 有关键字跨 run 拆分，段落级替换：全量文字放入第一个 run
    new_text = full_text.replace(find_text, replace_text)
    if runs:
        runs[0].text = new_text
        for run in runs[1:]:
            run.text = ""
    return occurrence_count


def replace_in_docx(filepath, find_text, replace_text):
    """对 .docx 文件执行替换，返回替换次数。
    先通过 python-docx 高层 API 处理正文/表格/页眉页脚，
    再通过 XML 级别扫描覆盖内容控件(SDT)、文本框等被遗漏的结构。"""
    try:
        from docx import Document
    except ImportError:
        return -1

    doc = Document(filepath)
    count = 0

    for para in doc.paragraphs:
        count += _replace_in_runs(para.runs, find_text, replace_text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    count += _replace_in_runs(para.runs, find_text, replace_text)
                # 嵌套表格
                for nested_table in cell.tables:
                    for nrow in nested_table.rows:
                        for ncell in nrow.cells:
                            for para in ncell.paragraphs:
                                count += _replace_in_runs(para.runs, find_text, replace_text)

    # 页眉页脚
    for section in doc.sections:
        for para in section.header.paragraphs:
            count += _replace_in_runs(para.runs, find_text, replace_text)
        for para in section.footer.paragraphs:
            count += _replace_in_runs(para.runs, find_text, replace_text)

    # XML 级别扫描：覆盖内容控件(SDT)、文本框等 python-docx 高层 API 遗漏的结构
    NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    def _xml_per_run_scan(root_elem):
        """逐个 w:t 元素替换，返回替换次数。"""
        n = 0
        for elem in root_elem.iter():
            if elem.tag == f'{{{NS}}}t' and elem.text and find_text in elem.text:
                n += elem.text.count(find_text)
                elem.text = elem.text.replace(find_text, replace_text)
        return n

    def _xml_cross_run_scan(root_elem):
        """跨 w:t 替换：处理文字被拆分到同一段落内多个 w:t 的情况。"""
        n = 0
        for para_elem in root_elem.iter(f'{{{NS}}}p'):
            wt_elems = para_elem.findall(f'.//{{{NS}}}t')
            if len(wt_elems) < 2:
                continue
            combined = ''.join(wt.text or '' for wt in wt_elems)
            if find_text in combined:
                n += combined.count(find_text)
                wt_elems[0].text = combined.replace(find_text, replace_text)
                for wt in wt_elems[1:]:
                    wt.text = ''
        return n

    count += _xml_per_run_scan(doc.element.body)
    count += _xml_cross_run_scan(doc.element.body)

    # 页眉页脚的 XML 级别扫描
    for section in doc.sections:
        for part in (section.header, section.footer):
            count += _xml_per_run_scan(part._element)
            count += _xml_cross_run_scan(part._element)

    if count:
        doc.save(filepath)
    return count


def replace_in_xlsx(filepath, find_text, replace_text):
    """对 Excel 文件执行替换，只处理字符串单元格，忽略公式，返回替换次数。"""
    try:
        from openpyxl import load_workbook
    except ImportError:
        return -1

    wb = load_workbook(filepath)
    count = 0

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                # 跳过公式单元格
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    continue
                if isinstance(cell.value, str) and find_text in cell.value:
                    cell.value = cell.value.replace(find_text, replace_text)
                    count += 1

    if count:
        wb.save(filepath)
    else:
        wb.close()
    return count


def replace_in_pptx(filepath, find_text, replace_text):
    """对 .pptx 文件执行替换，返回替换次数。"""
    try:
        from pptx import Presentation
    except ImportError:
        return -1

    prs = Presentation(filepath)
    count = 0

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    count += _replace_in_runs(para.runs, find_text, replace_text)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for para in cell.text_frame.paragraphs:
                            count += _replace_in_runs(para.runs, find_text, replace_text)

    if count:
        prs.save(filepath)
    return count


# ── COM 单例（避免反复创建/销毁导致挂起）──────────────────────────

_word = None
_excel = None


def _get_word():
    global _word
    if _word is None:
        import win32com.client
        _word = win32com.client.Dispatch("Word.Application")
        _word.Visible = False
        _word.DisplayAlerts = 0
        _word.ScreenUpdating = False
    return _word


def _get_excel():
    global _excel
    if _excel is None:
        import win32com.client
        _excel = win32com.client.Dispatch("Excel.Application")
        _excel.Visible = False
        _excel.DisplayAlerts = False
        _excel.ScreenUpdating = False
        _excel.EnableEvents = False
    return _excel


def _cleanup_com():
    global _word, _excel
    for obj in (_word, _excel):
        if obj is not None:
            try:
                obj.Quit()
            except Exception:
                pass
    _word = None
    _excel = None


def replace_in_doc(filepath, find_text, replace_text):
    """对 .doc 旧格式文件执行替换（通过 Word COM），返回替换次数。
    遇到路径包含特殊字符导致 Word 无法打开时，自动复制到临时目录处理。"""
    try:
        import win32com.client
    except ImportError:
        return -1

    filepath = str(Path(filepath).resolve())

    def _open_and_replace(path):
        doc = word.Documents.Open(path)
        try:
            original = doc.Content.Text
            if find_text not in original:
                return 0
            count = original.count(find_text)
            word.Selection.HomeKey(Unit=6)  # wdStory
            find_obj = word.Selection.Find
            find_obj.ClearFormatting()
            find_obj.Replacement.ClearFormatting()
            find_obj.Execute(
                find_text, False, False, False, False, False,
                True, 1, False, replace_text, 2,
            )
            doc.Save()
            return count
        finally:
            doc.Close()

    word = _get_word()

    # Try direct open first
    try:
        return _open_and_replace(filepath)
    except Exception:
        pass

    # Fallback: the path may contain chars Word COM can't handle
    # Copy to a temp dir with a simple ASCII name, process, copy back
    tmpdir = tempfile.mkdtemp()
    try:
        ext = os.path.splitext(filepath)[1]
        tmpfile = os.path.join(tmpdir, f"_tmp{ext}")
        shutil.copy2(filepath, tmpfile)
        count = _open_and_replace(tmpfile)
        if count:
            shutil.copy2(tmpfile, filepath)
        return count
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def replace_in_xls(filepath, find_text, replace_text):
    """对 .xls 旧格式文件执行替换，多层回退：Excel COM → xlrd+xlutils → xlrd+xlwt。"""

    # 方案一: Excel COM — 原生保留所有格式，最可靠
    try:
        import win32com.client
        excel = _get_excel()
        wb = excel.Workbooks.Open(filepath)
        if wb is not None:
            count = 0
            try:
                for ws in wb.Worksheets:
                    used = ws.UsedRange
                    if used is None:
                        continue
                    for row in used.Rows:
                        for cell in row.Cells:
                            if cell.Value is not None and isinstance(cell.Value, str) and find_text in cell.Value:
                                cell.Value = cell.Value.replace(find_text, replace_text)
                                count += 1
                if count:
                    wb.Save()
            finally:
                wb.Close()
            return count
    except ImportError:
        pass
    except Exception:
        pass

    # 方案二: xlrd + xlutils（保留格式，可能丢失被替换单元格的格式）
    try:
        import xlrd
        from xlutils.copy import copy
        rb = xlrd.open_workbook(filepath, formatting_info=True)
        wb = copy(rb)
        count = 0
        for sheet_idx in range(rb.nsheets):
            rs = rb.sheet_by_index(sheet_idx)
            ws = wb.get_sheet(sheet_idx)
            for row in range(rs.nrows):
                for col in range(rs.ncols):
                    try:
                        value = rs.cell_value(row, col)
                    except Exception:
                        continue
                    if isinstance(value, str) and find_text in value:
                        ws.write(row, col, value.replace(find_text, replace_text))
                        count += 1
        if count:
            wb.save(filepath)
        return count
    except ImportError:
        pass
    except Exception:
        pass

    # 方案三: xlrd 只读 + xlwt 重写（丢失全部格式，最后手段）
    try:
        import xlrd
        import xlwt
        rb = xlrd.open_workbook(filepath, formatting_info=False)
        wb = xlwt.Workbook()
        count = 0
        for sheet_idx in range(rb.nsheets):
            rs = rb.sheet_by_index(sheet_idx)
            ws = wb.add_sheet(rs.name or f"Sheet{sheet_idx+1}")
            for row in range(rs.nrows):
                for col in range(rs.ncols):
                    try:
                        value = rs.cell_value(row, col)
                    except Exception:
                        continue
                    if isinstance(value, str) and find_text in value:
                        value = value.replace(find_text, replace_text)
                        count += 1
                    if value != '' and value is not None:
                        ws.write(row, col, value)
        if count:
            wb.save(filepath)
        return count
    except ImportError:
        pass
    except Exception:
        pass

    return 0


def process_folder(folder, find_text, replace_text, log_callback=None):
    """
    遍历文件夹，对每个文件执行替换。
    返回 {"文件路径": 替换次数, ...}
    """
    results = {}
    total_replaced = 0

    for root, dirs, files in os.walk(folder):
        for filename in files:
            filepath = os.path.join(root, filename)
            ext = Path(filename).suffix.lower()

            if ext in SKIP_EXTENSIONS:
                continue

            if log_callback:
                log_callback(f"处理: {filepath}")

            try:
                if ext == ".docx":
                    cnt = replace_in_docx(filepath, find_text, replace_text)
                elif ext == ".doc":
                    cnt = replace_in_doc(filepath, find_text, replace_text)
                elif ext == ".xlsx":
                    cnt = replace_in_xlsx(filepath, find_text, replace_text)
                elif ext == ".xls":
                    cnt = replace_in_xls(filepath, find_text, replace_text)
                elif ext == ".pptx":
                    cnt = replace_in_pptx(filepath, find_text, replace_text)
                elif ext in TEXT_EXTENSIONS:
                    cnt = replace_in_text_file(filepath, find_text, replace_text)
                else:
                    # 未知类型当作文本尝试
                    cnt = replace_in_text_file(filepath, find_text, replace_text)

                if cnt is None:
                    cnt = 0
                if cnt == -1:
                    # 缺少依赖库
                    results[filepath] = "缺少依赖库"
                elif cnt > 0:
                    results[filepath] = cnt
                    total_replaced += cnt
            except Exception as e:
                results[filepath] = f"错误: {str(e)[:80]}"

    return results, total_replaced


# ── GUI ─────────────────────────────────────────────────────────────

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("批量文字替换工具")
        self.root.geometry("720x580")
        self.root.resizable(True, True)

        # 顶部区域 - 文件夹选择
        top = ttk.Frame(root, padding=12)
        top.pack(fill=tk.X)

        ttk.Label(top, text="源文件夹:").pack(side=tk.LEFT)
        self.folder_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.folder_var, width=55).pack(side=tk.LEFT, padx=6)
        ttk.Button(top, text="选择...", command=self.choose_folder).pack(side=tk.LEFT)

        # 查找 / 替换输入
        mid = ttk.Frame(root, padding=12)
        mid.pack(fill=tk.X)

        ttk.Label(mid, text="查找文字:").grid(row=0, column=0, sticky=tk.W, pady=4)
        self.find_var = tk.StringVar()
        ttk.Entry(mid, textvariable=self.find_var, width=30).grid(row=0, column=1, padx=6, pady=4, sticky=tk.W)

        ttk.Label(mid, text="替换为:").grid(row=1, column=0, sticky=tk.W, pady=4)
        self.replace_var = tk.StringVar()
        ttk.Entry(mid, textvariable=self.replace_var, width=30).grid(row=1, column=1, padx=6, pady=4, sticky=tk.W)

        # 执行按钮
        btn_frame = ttk.Frame(mid)
        btn_frame.grid(row=1, column=2, rowspan=2, padx=12)
        self.run_btn = ttk.Button(btn_frame, text="开始替换", command=self.start)
        self.run_btn.pack()

        # 进度条
        self.progress = ttk.Progressbar(root, mode="indeterminate")

        # 结果展示
        result_frame = ttk.Frame(root, padding=12)
        result_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(result_frame, text="替换报告:").pack(anchor=tk.W)
        self.result_text = tk.Text(result_frame, wrap=tk.WORD, font=("Consolas", 10))
        scroll = ttk.Scrollbar(result_frame, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scroll.set)
        self.result_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        scroll.pack(fill=tk.Y, side=tk.RIGHT)

        # 底部
        ttk.Label(root, text="提示: 操作会先复制源文件夹，在副本中替换，原文件不受影响。",
                  padding=6).pack()

    def choose_folder(self):
        path = filedialog.askdirectory(title="选择要处理的文件夹")
        if path:
            self.folder_var.set(path)

    def log(self, message):
        self.result_text.insert(tk.END, message + "\n")
        self.result_text.see(tk.END)
        self.root.update_idletasks()

    def start(self):
        src = self.folder_var.get().strip()
        find_text = self.find_var.get()
        replace_text = self.replace_var.get()

        if not src or not os.path.isdir(src):
            messagebox.showerror("错误", "请先选择一个有效的文件夹。")
            return
        if not find_text:
            messagebox.showerror("错误", "请输入要查找的文字。")
            return
        if find_text == replace_text:
            messagebox.showinfo("提示", "查找文字与替换文字相同，无需操作。")
            return

        self.run_btn.configure(state="disabled")
        self.result_text.delete("1.0", tk.END)
        self.progress.pack(fill=tk.X, before=self.result_text.master, pady=(0, 4))
        self.progress.start()

        threading.Thread(target=self._run, args=(src, find_text, replace_text),
                         daemon=True).start()

    def _run(self, src, find_text, replace_text):
        try:
            # 1. 确定目标文件夹名
            parent = os.path.dirname(src.rstrip(os.sep))
            src_name = os.path.basename(src.rstrip(os.sep))
            dst = os.path.join(parent, f"{src_name}_replaced")

            # 若已存在则加序号
            if os.path.exists(dst):
                i = 1
                while os.path.exists(f"{dst}_{i}"):
                    i += 1
                dst = f"{dst}_{i}"

            self.log(f"源文件夹: {src}")
            self.log(f"目标文件夹: {dst}")
            self.log(f"查找: \"{find_text}\"  →  替换: \"{replace_text}\"")
            self.log("=" * 60)

            # 2. 复制
            self.log("\n正在复制文件夹...")
            copy_folder(src, dst)
            self.log("复制完成。\n")

            # 3. 扫描替换
            self.log("正在扫描并替换...\n")
            results, total = process_folder(dst, find_text, replace_text,
                                            log_callback=self.log)

            # 4. 汇总报告
            self.log("\n" + "=" * 60)
            self.log("替换完成!\n")

            modified = {k: v for k, v in results.items() if isinstance(v, int) and v > 0}
            errors = {k: v for k, v in results.items() if isinstance(v, str)}

            if modified:
                self.log(f"修改文件数: {len(modified)} 个，共替换 {total} 处\n")
                for path, cnt in sorted(modified.items()):
                    rel = os.path.relpath(path, dst)
                    self.log(f"  {rel}  →  {cnt} 处")
            else:
                self.log("没有找到匹配的文字。")

            if errors:
                self.log(f"\n跳过/出错文件: {len(errors)} 个")
                for path, reason in sorted(errors.items()):
                    rel = os.path.relpath(path, dst)
                    self.log(f"  {rel}  →  {reason}")

        except Exception as e:
            self.log(f"\n发生错误: {e}")
        finally:
            _cleanup_com()
            self.root.after(0, self._finish)

    def _finish(self):
        self.progress.stop()
        self.progress.pack_forget()
        self.run_btn.configure(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
