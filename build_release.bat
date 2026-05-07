@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo 正在创建虚拟环境 .venv ...
  py -3 -m venv .venv
  if errorlevel 1 (
    python -m venv .venv
  )
)
call .venv\Scripts\activate.bat
python -m pip install -q -r requirements-dev.txt

python -m PyInstaller --noconfirm 智能排款.spec
if errorlevel 1 exit /b 1

if exist "dist\智能排款\AutoPaymentScheduleFile" rmdir /s /q "dist\智能排款\AutoPaymentScheduleFile"
xcopy /e /i /q "AutoPaymentScheduleFile" "dist\智能排款\AutoPaymentScheduleFile\" >nul

echo 完成。请将 dist\智能排款 整个文件夹发给对方（内含 智能排款.exe 与 AutoPaymentScheduleFile）。详见 打包与分发说明.txt。
