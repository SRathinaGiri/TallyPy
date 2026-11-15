@echo off
setlocal
call D:\hmg5\build64.bat /n connection_test.hbp
if errorlevel 1 exit /b %errorlevel%
if not exist connection_test.exe exit /b 1
exit /b 0
