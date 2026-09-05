@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

set "APP_NAME=Tally Qt Exporter"
if defined ProgramW6432 (
    set "INSTALL_DIR=%ProgramW6432%\Tally Qt Exporter"
) else (
    set "INSTALL_DIR=%ProgramFiles%\Tally Qt Exporter"
)
if not "%ProgramFiles(x86)%"=="" set "LEGACY_INSTALL_DIR=%ProgramFiles(x86)%\Tally Qt Exporter"
set "APP_FILES_DIR=%SCRIPT_DIR%\TallyQtExporter"
set "UNINSTALL_KEY=HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\TallyQtExporter"

if not exist "%APP_FILES_DIR%\TallyQtExporter.exe" (
    echo Package files not found: %APP_FILES_DIR%
    exit /b 1
)

if exist "%INSTALL_DIR%" rmdir /s /q "%INSTALL_DIR%"
if defined LEGACY_INSTALL_DIR if /I not "%LEGACY_INSTALL_DIR%"=="%INSTALL_DIR%" if exist "%LEGACY_INSTALL_DIR%" rmdir /s /q "%LEGACY_INSTALL_DIR%"
reg delete "HKLM\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\TallyQtExporter" /f >nul 2>nul
mkdir "%INSTALL_DIR%"

xcopy "%APP_FILES_DIR%\*" "%INSTALL_DIR%\" /E /I /Y >nul
if errorlevel 1 exit /b 1

set "TARGET_EXE=%INSTALL_DIR%\TallyQtExporter.exe"
set "UNINSTALL_CMD=%INSTALL_DIR%\uninstall.cmd"
set "POWERSHELL_EXE=%WINDIR%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%POWERSHELL_EXE%" set "POWERSHELL_EXE=%WINDIR%\System32\WindowsPowerShell\v1.0\powershell.exe"

"%POWERSHELL_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$key = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\TallyQtExporter'; New-Item -Path $key -Force | Out-Null; Set-ItemProperty -Path $key -Name DisplayName -Value $env:APP_NAME; Set-ItemProperty -Path $key -Name DisplayVersion -Value '0.1.10'; Set-ItemProperty -Path $key -Name Publisher -Value 'TallyXML'; Set-ItemProperty -Path $key -Name InstallLocation -Value $env:INSTALL_DIR; Set-ItemProperty -Path $key -Name DisplayIcon -Value $env:TARGET_EXE; Set-ItemProperty -Path $key -Name UninstallString -Value ([char]34 + $env:UNINSTALL_CMD + [char]34); New-ItemProperty -Path $key -Name NoModify -Value 1 -PropertyType DWord -Force | Out-Null; New-ItemProperty -Path $key -Name NoRepair -Value 1 -PropertyType DWord -Force | Out-Null"
if errorlevel 1 exit /b 1

"%POWERSHELL_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$WshShell = New-Object -ComObject WScript.Shell; $desktop = [Environment]::GetFolderPath('Desktop'); $startMenu = Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs'; $shortcut = $WshShell.CreateShortcut((Join-Path $desktop 'Tally Qt Exporter.lnk')); $shortcut.TargetPath = $env:TARGET_EXE; $shortcut.WorkingDirectory = $env:INSTALL_DIR; $shortcut.IconLocation = $env:TARGET_EXE + ',0'; $shortcut.Save(); $shortcut2 = $WshShell.CreateShortcut((Join-Path $startMenu 'Tally Qt Exporter.lnk')); $shortcut2.TargetPath = $env:TARGET_EXE; $shortcut2.WorkingDirectory = $env:INSTALL_DIR; $shortcut2.IconLocation = $env:TARGET_EXE + ',0'; $shortcut2.Save();"
if errorlevel 1 exit /b 1

start "" "%TARGET_EXE%"
exit /b 0
