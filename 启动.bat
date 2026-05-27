@echo off
chcp 65001 >nul

REM 寻找可用的 Python 解释器（优先 pythonw 无控制台窗口）
set PYTHON=

REM 1) 尝试 pythonw（python.org 发行版才有）
pythonw --version >nul 2>&1
if not errorlevel 1 (
    set PYTHON=pythonw
    goto :run
)

REM 2) 尝试 python
python --version >nul 2>&1
if not errorlevel 1 (
    set PYTHON=python
    goto :run
)

REM 3) 尝试 py 启动器 + pythonw
py -3w --version >nul 2>&1
if not errorlevel 1 (
    set PYTHON=py -3w
    goto :run
)

REM 4) 尝试 py 启动器 + python
py -3 --version >nul 2>&1
if not errorlevel 1 (
    set PYTHON=py -3
    goto :run
)

REM 全都没找到
echo [错误] 未找到 Python，请先运行 "安装依赖.bat"
pause
exit /b 1

:run
cd /d "%~dp0"
start "" %PYTHON% "%~dp0word_replace.py"
