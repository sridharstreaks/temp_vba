@echo off
REM --- Check if Python is installed ---
where python >nul 2>&1
if errorlevel 1 (
    echo Python is not installed. Please install Python and try again.
    pause
    exit /b 1
)

REM --- Defining paths ---
set "DESKTOP=%USERPROFILE%\Desktop"
set "DOWNLOADS=%USERPROFILE%\Downloads"
set "TARGET_DIR=%DESKTOP%\helium_chrome"

REM --- Creating the target folder on Desktop ---
if not exist "%TARGET_DIR%" (
    mkdir "%TARGET_DIR%"
    echo Created folder: %TARGET_DIR%
) else (
    echo Folder already exists: %TARGET_DIR%
)

REM --- Creating a virtual environment named ".venv" inside the target folder required for packages install---
cd /d "%TARGET_DIR%"
python -m venv .venv
if errorlevel 1 (
    echo Failed to create virtual environment.
    pause
    exit /b 1
)
echo Virtual environment created in %TARGET_DIR%\.venv

REM --- Activating the virtual environment (Windows activation) ---
call .venv\Scripts\activate.bat
if not defined VIRTUAL_ENV (
    echo Failed to activate virtual environment.
    pause
    exit /b 1
)
echo Virtual environment activated.

REM --- Upgrading pip to ensure Python is working ---
python.exe -m pip install --upgrade pip
if errorlevel 1 (
    echo pip upgrade failed.
    call deactivate
    pause
    exit /b 1
)
echo pip upgraded successfully.

REM --- Installing the 'helium' package ---
pip install helium
if errorlevel 1 (
    echo Failed to install the helium package.
    call deactivate
    pause
    exit /b 1
)
echo 'helium' package installed.

REM --- Searching for test.py in the Downloads folder ---
echo Searching for test.py in %DOWNLOADS%...
set "TEST_FILE="
for /r "%DOWNLOADS%" %%f in (test.py) do (
    set "TEST_FILE=%%f"
    goto FoundTest
)
:FoundTest
if "%TEST_FILE%"=="" (
    echo test.py not found in %DOWNLOADS%.
    call deactivate
    pause
    exit /b 1
)
echo Found test.py: %TEST_FILE%

REM --- Copy test.py to the helium_chrome folder ---
copy "%TEST_FILE%" "%TARGET_DIR%\test.py" >nul
echo Copied test.py to %TARGET_DIR%.

REM --- Run test.py using the Python interpreter from the virtual environment ---
echo Running test.py...
python "%TARGET_DIR%\test.py"
if errorlevel 1 (
    echo Failed to run test.py.
    call deactivate
    pause
    exit /b 1
)

REM --- Deactivate the virtual environment ---
call deactivate
echo Virtual environment deactivated.

pause