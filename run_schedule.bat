@echo off
chcp 65001 >nul
echo ========================================
echo 计划生成工具 - Schedule Generator
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到 Python，请先安装 Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] 检查依赖包...
pip install -q openpyxl pandas requests 2>nul

echo [2/3] 检查配置文件...
if not exist "config.ini" (
    echo 警告: 配置文件 config.ini 不存在
    echo 将使用默认配置
)

echo [3/3] 运行计划生成工具...
echo.

REM Run the main script
python call_ai.py

if errorlevel 1 (
    echo.
    echo 运行失败，请检查错误信息
    pause
    exit /b 1
)

echo.
echo ========================================
echo 运行完成！
echo ========================================
echo.
echo 生成的文件位置:
echo - Excel文件: output_transformed.xlsx
echo - AI响应: ai_responses/schedule_result.json
echo.
pause
