@echo off
REM -------------------------------
REM Launch OBS and automated Sipgate mic monitoring
REM -------------------------------

REM Try graceful close of OBS
taskkill /IM obs64.exe >nul 2>&1

REM Give OBS a few seconds to exit properly
timeout /t 3 /nobreak >nul

REM If OBS is still running, force kill
taskkill /IM obs64.exe /F >nul 2>&1

REM Kill your pythonw script if running (optional refinement below)
taskkill /IM pythonw.exe /F >nul 2>&1



REM Read JSON values with PowerShell

for /f "usebackq delims=" %%i in (`powershell -NoProfile -Command ^
  "(Get-Content 'config.json' | ConvertFrom-Json).obs_path"`) do set OBS_PATH=%%i

for /f "usebackq delims=" %%i in (`powershell -NoProfile -Command ^
  "(Get-Content 'config.json' | ConvertFrom-Json).obs_folder_path"`) do set OBS_FOLDER_PATH=%%i
  
for /f "usebackq delims=" %%i in (`powershell -NoProfile -Command ^
  "(Get-Content 'config.json' | ConvertFrom-Json).python_path"`) do set PYTHON_PATH=%%i


for /f "usebackq delims=" %%i in (`powershell -NoProfile -Command ^
  "(Get-Content 'config.json' | ConvertFrom-Json).project_path"`) do set PROJECT_PATH=%%i




REM Navigate to OBS folder
cd /d %OBS_FOLDER_PATH%

REM Start OBS
start "" %OBS_PATH% 

REM Wait a few seconds for OBS to launch
timeout /t 5 /nobreak

REM Navigate to your project folder
cd /d %PROJECT_PATH%

REM Start the automated mic-detection script
start "" "%PYTHON_PATH%" sipgate_mic_monitor.py

REM Wait a few seconds for python script to launch
timeout /t 5 /nobreak

REM Show reminder popup to check OBS sources
powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Reminder: Please check OBS Display Capture sources for your monitors and adjust your scene if needed .','OBS Setup Reminder','OK','Information')"

REM Done — exit silently
exit