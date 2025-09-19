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



REM Path to OBS executable (edit if needed)
set OBS_PATH="C:\Program Files\obs-studio\bin\64bit\obs64.exe"

REM Path to OBS folder (edit if needed)
set OBS_FOLDER_PATH="C:\Program Files\obs-studio\bin\64bit"

REM Path to virtual environment (venv) Python executable (edit if needed, 'pythonw' since windowless execution)
set PYTHON_PATH="C:\Users\LeonWeber\Leon\recording_tool\venv\Scripts\pythonw.exe"

REM Path to project folder (edit if needed)
set PROJECT_PATH="C:\Users\LeonWeber\Leon\recording_tool"





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

REM Show reminder popup to check OBS sources
powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Reminder: Please check OBS Display Capture sources for your monitors and adjust your scene if needed .','OBS Setup Reminder','OK','Information')"

REM Done — exit silently
exit