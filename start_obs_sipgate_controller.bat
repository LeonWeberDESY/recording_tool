@echo off
REM -------------------------------
REM Launch OBS and automated Sipgate mic monitoring
REM -------------------------------

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

REM Done — exit silently
exit