@echo off
setlocal

set "APP_NAME=Tally Qt Exporter"
if defined ProgramW6432 (
    set "INSTALL_DIR=%ProgramW6432%\Tally Qt Exporter"
) else (
    set "INSTALL_DIR=%ProgramFiles%\Tally Qt Exporter"
)
set "UNINSTALL_KEY=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\TallyQtExporter"

taskkill /IM TallyQtExporter.exe /F >nul 2>nul

del /f /q "%USERPROFILE%\Desktop\Tally Qt Exporter.lnk" >nul 2>nul
del /f /q "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Tally Qt Exporter.lnk" >nul 2>nul
reg delete "%UNINSTALL_KEY%" /f >nul 2>nul
reg delete "HKLM\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\TallyQtExporter" /f >nul 2>nul

if exist "%INSTALL_DIR%" (
    pushd "%TEMP%"
    start "" /min cmd /c "timeout /t 2 /nobreak >nul & rmdir /s /q \"%INSTALL_DIR%\""
    popd
)

exit /b 0
