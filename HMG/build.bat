@echo off
setlocal
call D:\hmg5\build64.bat /n TallyExporter.hbp
if errorlevel 1 exit /b %errorlevel%
if not exist TallyExporter.exe exit /b 1
exit /b 0

