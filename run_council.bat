@echo off
REM Starts Council without leaving a console window open.
cd /d "%~dp0"
where pyw >nul 2>nul && (start "" pyw council.py & exit /b)
where pythonw >nul 2>nul && (start "" pythonw council.py & exit /b)
python council.py
