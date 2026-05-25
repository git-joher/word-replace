@echo off
chcp 65001 >nul
echo ========================================
echo   批量文字替换工具 - 安装依赖
echo ========================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    echo 安装时请勾选 "Add Python to PATH"
    pause
    exit /b 1
)

echo 检测到 Python 环境，正在安装依赖...
echo.

pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [错误] 依赖安装失败，请检查网络连接后重试
    pause
    exit /b 1
)

echo.
echo ========================================
echo   安装完成！双击 "启动.bat" 启动程序
echo ========================================
pause
