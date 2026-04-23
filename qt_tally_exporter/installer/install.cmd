@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

set "APP_NAME=Tally Qt Exporter"
set "INSTALL_DIR=%ProgramFiles%\Tally Qt Exporter"
set "PAYLOAD_DIR=%SCRIPT_DIR%\payload"

if not exist "%PAYLOAD_DIR%\TallyQtExporter.exe" (
    echo Package payload not found: %PAYLOAD_DIR%
    exit /b 1
)

if exist "%INSTALL_DIR%" rmdir /s /q "%INSTALL_DIR%"
mkdir "%INSTALL_DIR%"

xcopy "%PAYLOAD_DIR%\*" "%INSTALL_DIR%\" /E /I /Y >nul
if errorlevel 1 exit /b 1

set "TARGET_EXE=%INSTALL_DIR%\TallyQtExporter.exe"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$WshShell = New-Object -ComObject WScript.Shell; $desktop = [Environment]::GetFolderPath('Desktop'); $startMenu = Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs'; $shortcut = $WshShell.CreateShortcut((Join-Path $desktop 'Tally Qt Exporter.lnk')); $shortcut.TargetPath = '%TARGET_EXE%'; $shortcut.WorkingDirectory = '%INSTALL_DIR%'; $shortcut.IconLocation = '%TARGET_EXE%,0'; $shortcut.Save(); $shortcut2 = $WshShell.CreateShortcut((Join-Path $startMenu 'Tally Qt Exporter.lnk')); $shortcut2.TargetPath = '%TARGET_EXE%'; $shortcut2.WorkingDirectory = '%INSTALL_DIR%'; $shortcut2.IconLocation = '%TARGET_EXE%,0'; $shortcut2.Save();"
if errorlevel 1 exit /b 1

start "" "%TARGET_EXE%"
exit /b 0
