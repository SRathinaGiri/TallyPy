@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "QT_ROOT=D:\Qt"
set "BUILD_DIR=%SCRIPT_DIR%\build"
set "DEPLOY_DIR=%SCRIPT_DIR%\dist\TallyQtExporter"
set "WINDEPLOYQT=%QT_ROOT%\6.11.0\mingw_64\bin\windeployqt.exe"
set "MINGW_BIN=%QT_ROOT%\Tools\mingw1310_64\bin"

if not exist "%BUILD_DIR%\TallyQtExporter.exe" (
    echo Release executable not found. Run build_release.bat first.
    exit /b 1
)

if not exist "%WINDEPLOYQT%" (
    echo windeployqt.exe not found at %WINDEPLOYQT%
    exit /b 1
)

if exist "%DEPLOY_DIR%" rmdir /s /q "%DEPLOY_DIR%"
mkdir "%DEPLOY_DIR%"

copy "%BUILD_DIR%\TallyQtExporter.exe" "%DEPLOY_DIR%\TallyQtExporter.exe" >nul
copy "%MINGW_BIN%\libgcc_s_seh-1.dll" "%DEPLOY_DIR%\" >nul
copy "%MINGW_BIN%\libstdc++-6.dll" "%DEPLOY_DIR%\" >nul
copy "%MINGW_BIN%\libwinpthread-1.dll" "%DEPLOY_DIR%\" >nul

call "%WINDEPLOYQT%" --release --compiler-runtime "%DEPLOY_DIR%\TallyQtExporter.exe"
if errorlevel 1 exit /b 1

echo Deployment complete: %DEPLOY_DIR%
