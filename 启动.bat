@echo off
chcp 65001 >nul

python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先运行 "安装依赖.bat"
    pause
    exit /b 1
)

cd /d "%~dp0"
start "" pythonw "%~dp0word_replace.py"
