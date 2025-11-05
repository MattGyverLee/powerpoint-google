@echo off
REM PowerPoint Cleaner - Windows Batch File
REM Drag and drop a .pptx file onto this batch file to clean it

echo ====================================
echo PowerPoint Cleaner for Google Slides
echo ====================================
echo.

if "%~1"=="" (
    echo Error: No file provided
    echo.
    echo Usage: Drag and drop a .pptx file onto this batch file
    echo    or: clean_powerpoint.bat "path\to\file.pptx"
    echo.
    pause
    exit /b 1
)

if not exist "%~1" (
    echo Error: File not found: %~1
    echo.
    pause
    exit /b 1
)

REM Get the directory where this batch file is located
set SCRIPT_DIR=%~dp0

REM Run the Python script
python "%SCRIPT_DIR%clean_powerpoint.py" "%~1"

echo.
echo ====================================
if %errorlevel% equ 0 (
    echo Success! Cleaned file created.
) else (
    echo An error occurred.
)
echo ====================================
echo.

pause
