@echo off
cd /d "%~dp0"
where pyw >nul 2>nul && ( start "" pyw "%~dp0gui.py" & exit /b )
where pythonw >nul 2>nul && ( start "" pythonw "%~dp0gui.py" & exit /b )
python "%~dp0gui.py"
