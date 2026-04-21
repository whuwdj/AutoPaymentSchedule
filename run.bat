@echo off
chcp 65001 >nul
cd /d "%~dp0"
if exist "%~dp0.venv\Scripts\pythonw.exe" (
  start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0main.py"
  exit /b 0
)
if exist "%~dp0.venv\Scripts\python.exe" (
  start "" "%~dp0.venv\Scripts\python.exe" "%~dp0main.py"
  exit /b 0
)
where pythonw >nul 2>&1
if %errorlevel%==0 (
  start "" pythonw "%~dp0main.py"
  exit /b 0
)
where python >nul 2>&1
if %errorlevel%==0 (
  start "" python "%~dp0main.py"
  exit /b 0
)
echo 未找到 Python，请先安装 Python 3.10+ 并勾选 "Add to PATH"。
pause
