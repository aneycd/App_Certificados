@echo off
REM Arranca la app con doble clic. Para detenerla: Ctrl+C en esta ventana.
cd /d "%~dp0"
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe" app.py
pause
