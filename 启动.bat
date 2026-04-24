@echo off
chcp 65001 >nul 2>&1
title 基站工参管理器

set "PYTHON=C:\Users\12931\.workbuddy\binaries\python\versions\3.13.12\python.exe"
set "PORT=18888"

if not exist "%PYTHON%" (
    echo [错误] 未找到 Python 环境: %PYTHON%
    echo 请联系管理员配置 Python 路径
    pause
    exit /b 1
)

cd /d "%~dp0"

:: 检测并释放端口占用
echo 正在检查端口 %PORT% ...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":%PORT% " ^| findstr "LISTENING"') do (
    echo 端口 %PORT% 被进程 %%a 占用，正在终止...
    taskkill /F /PID %%a >nul 2>&1
    timeout /t 1 /nobreak >nul
)
echo 端口 %PORT% 就绪
echo.

echo 正在启动基站工参管理器...
echo.

:: 记录当前进程PID，用于关闭时清理
for /f "tokens=2" %%p in ('tasklist /fi "imagename eq python.exe" /fo csv /nh ^| findstr /i "app.py"') do (
    set "APP_PID=%%~p"
)

:: 启动应用，Ctrl+C 或窗口关闭时触发清理
"%PYTHON%" "%~dp0app.py"
set "EXIT_CODE=%errorlevel%"

:: 清理：释放占用端口的进程
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":%PORT% " ^| findstr "LISTENING" 2^>nul') do (
    taskkill /F /PID %%a >nul 2>&1
)

if %EXIT_CODE% neq 0 (
    echo.
    echo [错误] 程序异常退出，请检查上方错误信息
    pause
)
