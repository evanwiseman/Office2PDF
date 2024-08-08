@echo off

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed. Please install Python 3.9 or higher.
    pause
    exit /b 1
)

REM Get Python version
for /f "tokens=2 delims= " %%G in ('python --version 2^>^&1') do (
    set "version=%%G"
)

REM Extract major and minor version numbers
for /f "tokens=1,2 delims=." %%G in ("%version%") do (
    set "major=%%G"
    set "minor=%%H"
)

REM Check if Python version is at least 3.9
if %major% lss 3 (
    echo Python version is less than 3.9. Please upgrade Python to 3.9 or higher.
    pause
    exit /b 1
)
if %major% equ 3 (
    if %minor% lss 9 (
        echo Python version is less than 3.9. Please upgrade Python to 3.9 or higher.
        pause
        exit /b 1
    )
)

REM Check if PyQt6 is installed
python -c "import PyQt6" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing PyQt6...
    python -m pip install PyQt6
)

REM Check if pywin32 is installed
python -c "import win32com" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing pywin32...
    python -m pip install pywin32
)

REM Check if Microsoft Word is installed
if exist "%ProgramFiles%\Microsoft Office\root\Office16\WINWORD.EXE" (
    echo Microsoft Word is installed.
) else if exist "%ProgramFiles(x86)%\Microsoft Office\root\Office16\WINWORD.EXE" (
    echo Microsoft Word is installed.
) else (
    echo Microsoft Word is not installed.
    set "word_installed=false"
)

REM Check if Microsoft Excel is installed
if exist "%ProgramFiles%\Microsoft Office\root\Office16\EXCEL.EXE" (
    echo Microsoft Excel is installed.
) else if exist "%ProgramFiles(x86)%\Microsoft Office\root\Office16\EXCEL.EXE" (
    echo Microsoft Excel is installed.
) else (
    echo Microsoft Excel is not installed.
    set "excel_installed=false"
)

REM Check if Microsoft PowerPoint is installed
if exist "%ProgramFiles%\Microsoft Office\root\Office16\POWERPNT.EXE" (
    echo Microsoft PowerPoint is installed.
) else if exist "%ProgramFiles(x86)%\Microsoft Office\root\Office16\POWERPNT.EXE" (
    echo Microsoft PowerPoint is installed.
) else (
    echo Microsoft PowerPoint is not installed.
    set "ppt_installed=false"
)

REM Check if all three are installed
if not "%word_installed%"=="false" (
    if not "%excel_installed%"=="false" (
        if not "%ppt_installed%"=="false" (
            echo All three Microsoft Office applications are installed.
        ) else (
            echo Microsoft PowerPoint is not installed.
            echo Please install Microsoft PowerPoint.
            pause
            exit /b 1
        )
    ) else (
        echo Microsoft Excel is not installed.
        echo Please install Microsoft Excel.
        pause
        exit /b 1
    )
) else (
    echo Microsoft Word is not installed.
    echo Please install Microsoft Word.
    pause
    exit /b 1
)

REM Run Application.py
python Application.py

pause