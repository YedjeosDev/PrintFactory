@echo off
setlocal
cd /d "%~dp0"

set "PYTHON_EXE=C:\Users\yedje\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if not exist "%PYTHON_EXE%" (
    set "PYTHON_EXE=python"
)

"%PYTHON_EXE%" -m pip install -r requirements.txt pyinstaller

"%PYTHON_EXE%" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name "GeneradorClientesEspeciales" ^
  --icon "Asset\icon.ico" ^
  --add-data "Logo;Logo" ^
  --add-data "C:\Users\yedje\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\Lib\tkinter;tkinter" ^
  --add-data "C:\Users\yedje\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\tcl;tcl" ^
  --add-binary "C:\Users\yedje\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\DLLs\_tkinter.pyd;." ^
  --add-binary "C:\Users\yedje\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\DLLs\tk86t.dll;." ^
  --add-binary "C:\Users\yedje\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\DLLs\tcl86t.dll;." ^
  --hidden-import _tkinter ^
  --hidden-import win32timezone ^
  --collect-all customtkinter ^
  --collect-all openpyxl ^
  --collect-all reportlab ^
  app.py

endlocal
