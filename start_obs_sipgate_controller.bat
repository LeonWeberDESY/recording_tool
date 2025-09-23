@echo off
REM -------------------------------
REM Launch OBS and automated Sipgate mic monitoring
REM -------------------------------

REM Set script directory as working directory
cd /d "%~dp0"

REM Check if config.json exists
if not exist "config.json" (
    echo ERROR: config.json not found in %~dp0
    echo Please ensure config.json is in the same directory as this script.
    pause
    exit /b 1
)

echo Stopping any running instances of OBS and the Python script...
REM Try graceful close of OBS
taskkill /IM obs64.exe >nul 2>&1

REM Give OBS a few seconds to exit properly
timeout /t 3 /nobreak >nul

REM If OBS is still running, force kill
taskkill /IM obs64.exe /F >nul 2>&1

REM Kill your pythonw script if running
taskkill /IM pythonw.exe /F >nul 2>&1

REM Add delay to prevent memory resource issues
timeout /t 2 /nobreak >nul

REM Read JSON values with PowerShell (with error handling)
echo Reading configuration...

for /f "usebackq delims=" %%i in (`powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { (Get-Content '%~dp0config.json' | ConvertFrom-Json).obs_path } catch { Write-Host 'ERROR_READING_CONFIG' }"`) do (
    if "%%i"=="ERROR_READING_CONFIG" (
        echo ERROR: Failed to read obs_path from config.json
        pause
        exit /b 1
    )
    set "OBS_PATH=%%i"
)

for /f "usebackq delims=" %%i in (`powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { (Get-Content '%~dp0config.json' | ConvertFrom-Json).obs_folder_path } catch { Write-Host 'ERROR_READING_CONFIG' }"`) do (
    if "%%i"=="ERROR_READING_CONFIG" (
        echo ERROR: Failed to read obs_folder_path from config.json
        pause
        exit /b 1
    )
    set "OBS_FOLDER_PATH=%%i"
)

for /f "usebackq delims=" %%i in (`powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { (Get-Content '%~dp0config.json' | ConvertFrom-Json).python_path } catch { Write-Host 'ERROR_READING_CONFIG' }"`) do (
    if "%%i"=="ERROR_READING_CONFIG" (
        echo ERROR: Failed to read python_path from config.json
        pause
        exit /b 1
    )
    set "PYTHON_PATH=%%i"
)

for /f "usebackq delims=" %%i in (`powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { (Get-Content '%~dp0config.json' | ConvertFrom-Json).project_path } catch { Write-Host 'ERROR_READING_CONFIG' }"`) do (
    if "%%i"=="ERROR_READING_CONFIG" (
        echo ERROR: Failed to read project_path from config.json
        pause
        exit /b 1
    )
    set "PROJECT_PATH=%%i"
)

REM Validate paths exist
if not exist "%OBS_PATH%" (
    echo ERROR: OBS executable not found at: %OBS_PATH%
    pause
    exit /b 1
)

if not exist "%OBS_FOLDER_PATH%" (
    echo ERROR: OBS folder not found at: %OBS_FOLDER_PATH%
    pause
    exit /b 1
)

if not exist "%PYTHON_PATH%" (
    echo ERROR: Python executable not found at: %PYTHON_PATH%
    pause
    exit /b 1
)

if not exist "%PROJECT_PATH%" (
    echo ERROR: Project folder not found at: %PROJECT_PATH%
    pause
    exit /b 1
)

if not exist "%PROJECT_PATH%\sipgate_mic_monitor.py" (
    echo ERROR: Python script not found at: %PROJECT_PATH%\sipgate_mic_monitor.py
    pause
    exit /b 1
)

echo Starting OBS...
REM Navigate to OBS folder
cd /d "%OBS_FOLDER_PATH%"

REM Start OBS (minimized to reduce resource impact)
start /MIN "" "%OBS_PATH%"

REM Wait for OBS to launch
echo Waiting for OBS to start...
timeout /t 8 /nobreak >nul

echo Starting Python monitoring script...
REM Navigate to your project folder
cd /d "%PROJECT_PATH%"

REM Start the automated mic-detection script (hidden window)
start /MIN "" "%PYTHON_PATH%" sipgate_mic_monitor.py

REM Wait for python script to launch
timeout /t 3 /nobreak >nul

REM Show reminder popup (with shorter message for Task Scheduler)
powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Reminder: Please check OBS Display Capture sources for your monitors and adjust your scene if needed .','OBS Setup Reminder','OK','Information')" >nul 2>&1


echo Setup complete!
REM Exit cleanly
exit /b 0