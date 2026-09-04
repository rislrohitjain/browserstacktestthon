@echo off
setlocal enabledelayedexpansion

echo ==============================================================================
echo        BrowserStack Testathon - Environment Setup and Installation
echo ==============================================================================
echo.

:: 1. Verify Python availability
echo [1/5] Checking Python installation...
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Python is not installed or not added to system PATH.
    echo Please install Python 3.10+ from https://www.python.org/ and retry.
    pause
    exit /b 1
)
for /f "tokens=*" %%i in ('python --version') do set PYTHON_VER=%%i
echo   - Found: !PYTHON_VER!
echo.

:: 2. Setup project-isolated virtual environment
echo [2/5] Initializing isolated virtual environment (./env)...
if exist "env\Scripts\activate.bat" (
    echo   - Virtual environment ./env already exists. Reusing existing environment.
) else (
    echo   - Creating new virtual environment at .\env ...
    python -m venv env
    if %ERRORLEVEL% neq 0 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo   - Virtual environment created successfully.
)
echo.

:: 3. Activate virtual environment
echo [3/5] Activating virtual environment...
call .\env\Scripts\activate.bat
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Failed to activate virtual environment.
    pause
    exit /b 1
)
echo   - Virtual environment active: %VIRTUAL_ENV%
echo.

:: 4. Upgrade pip and install dependencies
echo [4/5] Upgrading pip and installing requirements...
python -m pip install --upgrade pip
if %ERRORLEVEL% neq 0 (
    echo [WARNING] Failed to upgrade pip. Continuing with dependency installation...
)

echo   - Installing packages from requirements.txt...
python -m pip install -r requirements.txt
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Failed to install one or more dependencies from requirements.txt.
    pause
    exit /b 1
)
echo   - All dependencies installed successfully.
echo.

:: 5. Check environment configuration (.env)
echo [5/5] Checking environment configuration...
if not exist ".env" (
    if exist ".env.example" (
        echo   - No .env file found. Copying .env.example to .env ...
        copy .env.example .env >nul
        echo   - Created .env template.
        echo   [ACTION REQUIRED] Please open .env and provide your BrowserStack and Database credentials!
    ) else (
        echo   [WARNING] Neither .env nor .env.example was found.
    )
) else (
    echo   - Found existing .env file.
)
echo.

echo ==============================================================================
echo   Setup completed successfully!
echo.
echo   To execute the test suite:
echo     1. Activate environment:   call .\env\Scripts\activate.bat
echo     2. Run test runner:        python test_runner.py
echo.
echo   Advanced flags:
echo     python test_runner.py --help
echo ==============================================================================
if /i "%~1" neq "/nopause" if /i "%~1" neq "--nopause" pause
