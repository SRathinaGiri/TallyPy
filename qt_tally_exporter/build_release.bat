@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "QT_ROOT=D:\Qt"
set "CMAKE_BIN=%QT_ROOT%\Tools\CMake_64\bin"
set "NINJA_BIN=%QT_ROOT%\Tools\Ninja"
set "MINGW_BIN=%QT_ROOT%\Tools\mingw1310_64\bin"
set "QT_PREFIX=%QT_ROOT%\6.11.0\mingw_64"
set "BUILD_DIR=%SCRIPT_DIR%\build"

set "PATH=%CMAKE_BIN%;%NINJA_BIN%;%MINGW_BIN%;%PATH%"

if not exist "%QT_PREFIX%\lib\cmake\Qt6\Qt6Config.cmake" (
    echo Qt6Config.cmake not found under %QT_PREFIX%
    exit /b 1
)

if not exist "%BUILD_DIR%" mkdir "%BUILD_DIR%"

cmake -S "%SCRIPT_DIR%" -B "%BUILD_DIR%" -G Ninja -DCMAKE_BUILD_TYPE=Release -DCMAKE_PREFIX_PATH="%QT_PREFIX%"
if errorlevel 1 exit /b 1

cmake --build "%BUILD_DIR%" --config Release
if errorlevel 1 exit /b 1

echo Build complete: %BUILD_DIR%\TallyQtExporter.exe
