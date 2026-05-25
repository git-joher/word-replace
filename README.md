# 批量文字替换工具

对文件夹内所有可编辑文件进行批量文字查找替换。先复制源文件夹，在副本上操作，**原文件不受影响**。

## 支持格式

| 格式 | 扩展名 | 说明 |
|------|--------|------|
| Word | `.doc` `.docx` | 保留表格、格式、页眉页脚 |
| Excel | `.xls` `.xlsx` | 保留单元格格式，跳过公式 |
| PPT | `.pptx` | 保留格式 |
| 纯文本 | `.txt` `.csv` `.json` `.xml` `.html` `.py` `.js` 等 |

## 系统要求

- Windows 操作系统
- Microsoft Office（Word + Excel），用于处理 `.doc` 和 `.xls` 文件
- Python 3.8+（如果用 bat 启动）

## 快速开始

### 方式一：bat 脚本启动

1. 安装 Python 3.8+（勾选 "Add Python to PATH"）
2. 双击 `安装依赖.bat` 安装依赖库
3. 双击 `启动.bat` 启动程序

### 方式二：打包为 exe（无需安装 Python）

```
pip install pyinstaller
pyinstaller --onefile --windowed --name "批量文字替换工具" word_replace.py
```

打包后在 `dist` 目录中找到 exe，复制到任意位置即可双击运行。

## 使用方法

1. 点击 **选择...** 选择要处理的文件夹
2. 输入 **查找文字** 和 **替换为**
3. 点击 **开始替换**

程序会在源文件夹同级目录创建 `原文件夹名_replaced` 副本，所有替换在副本中进行。

## 依赖库

```
python-docx>=0.8.11
openpyxl>=3.0.0
python-pptx>=0.6.21
pywin32>=300
xlrd==1.2.0
xlutils>=2.0.0
xlwt>=1.3.0
```
