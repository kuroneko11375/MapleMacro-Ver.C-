@echo off
echo ========================================
echo  Building macro_core.dll (Rust ¡÷ x64)
echo ========================================

cd /d "%~dp0"

:: Check if cargo is available
where cargo >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo ERROR: Rust/Cargo not found!
    echo Please install Rust from https://rustup.rs/
    echo After installing, run this script again.
    pause
    exit /b 1
)

:: Build release
cargo build --release --target x86_64-pc-windows-msvc
if %ERRORLEVEL% neq 0 (
    echo ERROR: Build failed!
    pause
    exit /b 1
)

:: Copy DLL to C# output directories
set "DLL_SRC=target\x86_64-pc-windows-msvc\release\macro_core.dll"
set "DLL_DBG=..\MapleStoryMacro\bin\x64\Debug\net8.0-windows\"
set "DLL_REL=..\MapleStoryMacro\bin\x64\Release\net8.0-windows\"

if not exist "%DLL_DBG%" mkdir "%DLL_DBG%"
copy /Y "%DLL_SRC%" "%DLL_DBG%"

if not exist "%DLL_REL%" mkdir "%DLL_REL%"
copy /Y "%DLL_SRC%" "%DLL_REL%"

echo.
echo macro_core.dll built and copied successfully!
echo    Debug:   %DLL_DBG%macro_core.dll
echo    Release: %DLL_REL%macro_core.dll
echo.
pause
