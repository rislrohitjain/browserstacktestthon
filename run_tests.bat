@echo off
setlocal enabledelayedexpansion

echo ==============================================================================
echo        BrowserStack Testathon - One-Click Automated Test Runner
echo ==============================================================================
echo.

:: 1. Check if virtual environment exists
if not exist "env\Scripts\python.exe" (
    echo [ERROR] Virtual environment not found.
    echo Running setup.bat to initialize environment and install packages...
    call setup.bat
    if %ERRORLEVEL% neq 0 (
        echo [ERROR] Setup failed.
        pause
        exit /b 1
    )
)

:: 2. Activate virtual environment
call .\env\Scripts\activate.bat

:: 3. Run full testathon matrix
echo Starting full BrowserStack Testathon suite...
echo.
python test_runner.py --all

echo.
echo ==============================================================================
echo Starting BStackDemo E-Commerce E2E suite...
echo.
python bstackdemo_suite.py

echo.
echo ==============================================================================
echo   All tests executed!
echo   Excel and PDF reports saved in: .\reports\
echo ==============================================================================
echo.
pause
