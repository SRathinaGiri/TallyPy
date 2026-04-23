@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

set "DIST_DIR=%SCRIPT_DIR%\dist\TallyQtExporter"
set "INSTALLER_DIR=%SCRIPT_DIR%\installer"
set "STAGE_DIR=%INSTALLER_DIR%\stage"
set "OUTPUT_EXE=%SCRIPT_DIR%\dist\TallyQtExporter_Setup.exe"
set "ARCHIVE_FILE=%INSTALLER_DIR%\package.7z"
set "CONFIG_FILE=%INSTALLER_DIR%\config.txt"
set "SEVENZIP_EXE=C:\Program Files\7-Zip\7z.exe"
set "SEVENZIP_SFX=C:\Program Files\7-Zip\7z.sfx"

if not exist "%DIST_DIR%\TallyQtExporter.exe" (
    echo Portable deployment not found. Run deploy_release.bat first.
    exit /b 1
)

if not exist "%SEVENZIP_EXE%" (
    echo 7z.exe not found at %SEVENZIP_EXE%
    exit /b 1
)

if not exist "%SEVENZIP_SFX%" (
    echo 7z.sfx not found at %SEVENZIP_SFX%
    exit /b 1
)

taskkill /IM TallyQtExporter.exe /F >nul 2>nul

if exist "%STAGE_DIR%" rmdir /s /q "%STAGE_DIR%"
mkdir "%STAGE_DIR%"
mkdir "%STAGE_DIR%\payload"

xcopy "%DIST_DIR%\*" "%STAGE_DIR%\payload\" /E /I /Y >nul
if errorlevel 1 exit /b 1

copy "%INSTALLER_DIR%\install.cmd" "%STAGE_DIR%\install.cmd" >nul

(
echo ;!@Install@!UTF-8!
echo Title="Tally Qt Exporter Setup"
echo RunProgram="install.cmd"
echo GUIMode="2"
echo ;!@InstallEnd@!
) > "%CONFIG_FILE%"

if exist "%ARCHIVE_FILE%" del /f /q "%ARCHIVE_FILE%"
if exist "%OUTPUT_EXE%" del /f /q "%OUTPUT_EXE%"

pushd "%STAGE_DIR%"
"%SEVENZIP_EXE%" a -t7z "%ARCHIVE_FILE%" * -mx=9
if errorlevel 1 (
    popd
    exit /b 1
)
popd

copy /b "%SEVENZIP_SFX%" + "%CONFIG_FILE%" + "%ARCHIVE_FILE%" "%OUTPUT_EXE%" >nul
if errorlevel 1 exit /b 1

echo Installer created: %OUTPUT_EXE%
