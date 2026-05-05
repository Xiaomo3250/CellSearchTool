@echo off
chcp 65001 >nul 2>&1
title 基站工参管理器

:: ============================================
:: Python 路径：如果 python 已加入 PATH，可直接写 "python"
:: ============================================
set "PYTHON=C:\Users\12931\AppData\Local\Python\pythoncore-3.14-64\python.exe"
set "PORT=18888"

cd /d "%~dp0"

:: ============================================
:: 第一步：检查 Python 是否真正可用
:: ============================================
echo Checking Python environment...

:: 1a. 文件是否存在
if not exist "%PYTHON%" (
    echo [ERROR] Python not found at: %PYTHON%
    echo.
    echo Please install Python 3.8+ and set PYTHON variable in this script.
    echo.
    echo   Option 1: https://www.python.org/downloads/
    echo   Option 2: winget install Python.Python.3.12
    echo.
    echo If already installed, edit PYTHON path in this bat file:
    echo   e.g. set "PYTHON=C:\Users\%USERNAME%\AppData\Local\Programs\Python\Python312\python.exe"
    pause
    exit /b 1
)

:: 1b. 验证能真正执行（排除 Windows Store 存根）
"%PYTHON%" --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python exists but cannot execute properly.
    echo.
    echo This usually means the python command is a Microsoft Store stub,
    echo not a real Python interpreter.
    echo.
    echo Fix:
    echo   1. Open Settings ^> Apps ^> App execution aliases
    echo   2. Turn OFF python.exe and python3.exe
    echo   3. Install real Python from https://www.python.org/downloads/
    echo   4. Check "Add Python to PATH" during installation
    echo.
    echo Or run: winget install Python.Python.3.12
    pause
    exit /b 1
)

for /f "tokens=2" %%v in ('"%PYTHON%" --version 2^>^&1') do set "PY_VER=%%v"
echo   Python version: %PY_VER% [OK]

:: 1c. 检查关键依赖
"%PYTHON%" -c "import pandas, openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo [WARNING] Missing dependencies, installing...
    echo.
    "%PYTHON%" -m pip install pandas openpyxl
    if %errorlevel% neq 0 (
        echo.
        echo [ERROR] Dependency installation failed. Run manually:
        echo   pip install pandas openpyxl
        pause
        exit /b 1
    )
    echo  Dependencies installed [OK]
) else (
    echo  Dependencies check [OK]
)
echo.

:: ============================================
:: 第二步：检测并释放端口占用
:: ============================================
echo Checking port %PORT% ...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":%PORT% " ^| findstr "LISTENING"') do (
    echo  Port %PORT% occupied by PID %%a, terminating...
    taskkill /F /PID %%a >nul 2>&1
    timeout /t 1 /nobreak >nul
)
echo  Port %PORT% ready [OK]
echo.

:: ============================================
:: 第三步：启动应用
:: ============================================
echo Starting CellSearchTool...
echo.

"%PYTHON%" "%~dp0app.py"
set "EXIT_CODE=%errorlevel%"

:: 清理：释放占用端口的进程
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":%PORT% " ^| findstr "LISTENING" 2^>nul') do (
    taskkill /F /PID %%a >nul 2>&1
)

if %EXIT_CODE% neq 0 (
    echo.
    echo [ERROR] Program exited abnormally (code: %EXIT_CODE%)
    pause
)
