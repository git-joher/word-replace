# 多对文字替换 — 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将替换工具从单对输入改为动态行列表，支持一次执行多对顺序替换。

**Architecture:** 新增 `PairListFrame` 类管理行列表 UI，新增 `_replace_single` 辅助函数消除重复代码，修改 `process_folder` 接受 pairs 列表并逐对处理每个文件，修改 `App` 使用 `PairListFrame` 替换原有单行输入框。

**Tech Stack:** Python 3, tkinter, python-docx, openpyxl, python-pptx, pywin32

**Modifies:** `word_replace.py` (仅此一个文件)

---

## 文件结构

```
word_replace.py          ← 唯一修改的文件
  + _replace_single()    ← 新增：提取文件类型分发逻辑
  ~ process_folder()     ← 修改：接受 pairs 列表参数
  + PairListFrame        ← 新增：行列表 UI 组件
  ~ App                  ← 修改：使用 PairListFrame + pairs
```

---

### Task 1: 新增 `_replace_single` 辅助函数

**文件:** `word_replace.py`

将 `process_folder` 中的文件类型分发逻辑提取为独立函数，供多对替换循环复用。

- [ ] **Step 1: 在 `process_folder` 上方插入新函数**

在第 413 行（`def process_folder` 之前）插入：

```python
def _replace_single(filepath, ext, find_text, replace_text):
    """对单个文件执行一对替换。返回替换次数，-1 表示缺依赖，异常向上抛出。"""
    if ext == ".docx":
        return replace_in_docx(filepath, find_text, replace_text)
    elif ext == ".doc":
        return replace_in_doc(filepath, find_text, replace_text)
    elif ext == ".xlsx":
        return replace_in_xlsx(filepath, find_text, replace_text)
    elif ext == ".xls":
        return replace_in_xls(filepath, find_text, replace_text)
    elif ext == ".pptx":
        return replace_in_pptx(filepath, find_text, replace_text)
    elif ext in TEXT_EXTENSIONS:
        return replace_in_text_file(filepath, find_text, replace_text)
    else:
        return replace_in_text_file(filepath, find_text, replace_text)
```

- [ ] **Step 2: 验证语法**

```bash
python -c "import py_compile; py_compile.compile('word_replace.py', doraise=True); print('OK')"
```

---

### Task 2: 修改 `process_folder` 接受 pairs 列表

**文件:** `word_replace.py` 第 414–461 行

- [ ] **Step 1: 替换 `process_folder` 函数**

```python
def process_folder(folder, pairs, log_callback=None):
    """
    遍历文件夹，对每个文件依次执行多对替换（按 pairs 顺序）。
    pairs: [(find_text, replace_text), ...]
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

            file_count = 0
            file_error = None

            for find_text, replace_text in pairs:
                if find_text == replace_text or not find_text:
                    continue

                if log_callback:
                    log_callback(f"  [{find_text}  →  {replace_text}]")

                try:
                    cnt = _replace_single(filepath, ext, find_text, replace_text)
                    if cnt is None:
                        cnt = 0
                    if cnt == -1:
                        file_error = "缺少依赖库"
                    elif cnt > 0:
                        file_count += cnt
                except Exception as e:
                    file_error = f"错误: {str(e)[:80]}"
                    break

            if file_error:
                results[filepath] = file_error
            elif file_count > 0:
                results[filepath] = file_count
                total_replaced += file_count

    return results, total_replaced
```

- [ ] **Step 2: 验证语法**

```bash
python -c "import py_compile; py_compile.compile('word_replace.py', doraise=True); print('OK')"
```

---

### Task 3: 新增 `PairListFrame` 类

**文件:** `word_replace.py`，在 `App` 类之前插入（约第 464 行前）

- [ ] **Step 1: 插入完整的 `PairListFrame` 类**

```python
class PairListFrame(ttk.Frame):
    """多对替换输入的行列表 UI 组件。"""

    def __init__(self, parent):
        super().__init__(parent)
        self._rows = []  # [(find_var, replace_var)]

        # 表头
        header = ttk.Frame(self)
        header.pack(fill=tk.X, pady=(0, 4))
        ttk.Label(header, text="查找文字", width=28).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Label(header, text="", width=2).pack(side=tk.LEFT)
        ttk.Label(header, text="替换为", width=28).pack(side=tk.LEFT, padx=(4, 0))

        self._inner = None
        self.add_row()
        self._rebuild()

    # ── 公开接口 ──

    def add_row(self, find_text="", replace_text=""):
        self._rows.append((tk.StringVar(value=find_text), tk.StringVar(value=replace_text)))
        self._rebuild()

    def remove_row(self, index):
        if len(self._rows) <= 1:
            return
        self._rows.pop(index)
        self._rebuild()

    def move_up(self, index):
        if index <= 0:
            return
        self._rows[index], self._rows[index - 1] = self._rows[index - 1], self._rows[index]
        self._rebuild()

    def move_down(self, index):
        if index >= len(self._rows) - 1:
            return
        self._rows[index], self._rows[index + 1] = self._rows[index + 1], self._rows[index]
        self._rebuild()

    def get_pairs(self):
        """返回 [(find, replace), ...]，跳过空查找。"""
        pairs = []
        for fv, rv in self._rows:
            f = fv.get().strip()
            if f:
                pairs.append((f, rv.get()))
        return pairs

    # ── 内部 ──

    def _rebuild(self):
        if self._inner is not None:
            self._inner.destroy()
        self._inner = ttk.Frame(self)
        self._inner.pack(fill=tk.X)

        for i, (fv, rv) in enumerate(self._rows):
            row_frame = ttk.Frame(self._inner)
            row_frame.pack(fill=tk.X, pady=1)

            ttk.Entry(row_frame, textvariable=fv, width=30).pack(side=tk.LEFT, padx=(0, 4))
            ttk.Label(row_frame, text="→").pack(side=tk.LEFT)
            ttk.Entry(row_frame, textvariable=rv, width=30).pack(side=tk.LEFT, padx=(4, 6))

            if len(self._rows) > 1:
                ttk.Button(row_frame, text="↑", width=2,
                           command=lambda idx=i: self.move_up(idx)).pack(side=tk.LEFT, padx=1)
                ttk.Button(row_frame, text="↓", width=2,
                           command=lambda idx=i: self.move_down(idx)).pack(side=tk.LEFT, padx=1)
                ttk.Button(row_frame, text="✕", width=2,
                           command=lambda idx=i: self.remove_row(idx)).pack(side=tk.LEFT, padx=1)

        # 添加按钮
        ttk.Button(self._inner, text="+ 添加替换对", command=self.add_row).pack(
            fill=tk.X, pady=(6, 0))
```

- [ ] **Step 2: 验证语法**

```bash
python -c "import py_compile; py_compile.compile('word_replace.py', doraise=True); print('OK')"
```

---

### Task 4: 修改 `App` 类

**文件:** `word_replace.py` 第 466–610 行

- [ ] **Step 1: 修改 `__init__` — 替换单行输入框为 PairListFrame**

将第 482–498 行（mid 框架内从 ttk.Label "查找文字" 到 run_btn）替换为：

```python
        # 多对替换输入
        mid = ttk.Frame(root, padding=12)
        mid.pack(fill=tk.X)

        self.pair_list = PairListFrame(mid)
        self.pair_list.pack(fill=tk.X)

        # 执行按钮
        self.run_btn = ttk.Button(mid, text="开始替换", command=self.start)
        self.run_btn.pack(pady=(12, 0))
```

- [ ] **Step 2: 修改 `start` — 校验改为 pairs 列表**

将第 528–549 行替换为：

```python
    def start(self):
        src = self.folder_var.get().strip()
        pairs = self.pair_list.get_pairs()

        if not src or not os.path.isdir(src):
            messagebox.showerror("错误", "请先选择一个有效的文件夹。")
            return
        if not pairs:
            messagebox.showerror("错误", "请至少添加一对待替换的文字。")
            return

        self.run_btn.configure(state="disabled")
        self.result_text.delete("1.0", tk.END)
        self.progress.pack(fill=tk.X, before=self.result_text.master, pady=(0, 4))
        self.progress.start()

        threading.Thread(target=self._run, args=(src, pairs),
                         daemon=True).start()
```

- [ ] **Step 3: 修改 `_run` — 接受 pairs 并显示信息**

将第 551–553 行的方法签名和日志部分替换为：

```python
    def _run(self, src, pairs):
        try:
            parent = os.path.dirname(src.rstrip(os.sep))
            src_name = os.path.basename(src.rstrip(os.sep))
            dst = os.path.join(parent, f"{src_name}_replaced")

            if os.path.exists(dst):
                i = 1
                while os.path.exists(f"{dst}_{i}"):
                    i += 1
                dst = f"{dst}_{i}"

            self.log(f"源文件夹: {src}")
            self.log(f"目标文件夹: {dst}")
            self.log(f"共 {len(pairs)} 对替换:")
            for i, (f, r) in enumerate(pairs, 1):
                self.log(f"  {i}. \"{f}\"  →  \"{r}\"")
            self.log("=" * 60)

            self.log("\n正在复制文件夹...")
            copy_folder(src, dst)
            self.log("复制完成。\n")

            self.log("正在扫描并替换...\n")
            results, total = process_folder(dst, pairs, log_callback=self.log)
```

保留第 580–605 行的汇总报告和 finally 代码不变。

- [ ] **Step 4: 增大默认窗口高度**

第 470 行，将 `"720x580"` 改为 `"720x650"` 以容纳多行输入。

- [ ] **Step 5: 验证语法**

```bash
python -c "import py_compile; py_compile.compile('word_replace.py', doraise=True); print('OK')"
```

---

### Task 5: 重建 exe 并验证

- [ ] **Step 1: 清理旧 build 文件**

```bash
rm -rf "D:/project/word-replace/build" "D:/project/word-replace/"*.spec
```

- [ ] **Step 2: 构建 exe**

```bash
cd "D:/project/word-replace" && pyinstaller --onefile --windowed --clean --name "批量文字替换工具" --hidden-import win32com.client --hidden-import docx --hidden-import openpyxl --hidden-import pptx --hidden-import xlrd --hidden-import xlutils --hidden-import xlwt word_replace.py 2>&1 | tail -5
```

- [ ] **Step 3: 清理 build 文件**

```bash
rm -rf "D:/project/word-replace/build" "D:/project/word-replace/"*.spec
```

- [ ] **Step 4: 手动测试**

启动 exe，添加 2-3 对替换，选择一个测试文件夹，验证多对替换顺序执行正确。
