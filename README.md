# Automated OBS + Sipgate Recording Tool for Windows

This project provides a fully automated workflow to record Desktop Screens during **Sipgate** calls using **OBS Studio** on **Windows**. The microphone input is dynamically enabled only during calls, preventing constant Kuando status light triggers.

## Features

- Automatically monitors Sipgate microphone sessions.
- Dynamically creates/disables OBS mic input for recording.
- Uses OBS WebSocket (v5) to control recording.
- No manual hotkeys required.
- Video file gets safed in folder and renamed to Ticket ID (manual input after recording stopped)

## Requirements

### Software
| Software      | Minimum Version / Notes                        | Installation Link |
|---------------|-----------------------------------------------|-----------------|
| **OBS Studio** | ≥ 31.1.2 (WebSocket ≥ 5.x included)          | [https://obsproject.com/](https://obsproject.com/) |
| **Python**     | ≥ 3.13                                        | [https://www.python.org/downloads/](https://www.python.org/downloads/) |
| **Git** | ≥ 2.51.0 | [https://git-scm.com/downloads/win](https://git-scm.com/downloads/win) |

Optional for testing or development:
- VSCode or any Python IDE.

### Python Packages
| **Package**	| **Purpose** |
| ----- | ----- |
| **simpleobsws** |	OBS WebSocket client (v5 compatible) |
| **pycaw**	| Windows audio session monitoring |
| **comtypes** |	Required by pycaw |
| **psutil** |	Process/session management |


### Project files
| **File**	| **Purpose** |
| ----- | ----- |
| `.gitattributes` | Contains the file formatting regarding line endings (CLRF or LF). |
| `.gitignore` | Contains the files or folders that will ignored by Git when pushing or pulling. |
| `config.json` | Contains all configuration settings. |
| `install.sh` | Bash script to create virtual environment and install required packages. |
| `obs_control.py` |	Dynamically creates mic input and controls OBS recording. |
| `README.md` |	Contains this README file for installation and documentation. |
| `requirements.txt` |	Contains the required packages to be installed. |
| `sipgate_mic_monitor.py`	| Monitors Sipgate call sessions and triggers OBS control. |
| `start_obs_sipgate_controller.bat` |	Batch script to launch OBS + Python scripts. |

## Setup Instructions

### 1. Install Software
- Make sure that all installations are made in system paths that DO NOT contain spaces in their name (like "Program Files")
- Download and install **OBS Studio** (≥ 31.1.2) and ensure WebSocket ≥ 5.x is included.
    1. Open OBS Studio and decline the "auto-configuration wizard".
    2. Open OBS Studio → Tools (top bar) → WebSocket Server Settings.
    3. Check "Eanble WebSocket server", set port (default 4455), and optionally a password.
    4. This will open Windows Firewall Settings → Accept it.
    5. Create a new scene and name it i.e. `sipgate_scene`
    6. Go to Settings (bottom right) → Video
    7. Set Base (Canvas) Resolution and Output (Scaled) resolution each to `5360x1440`. Set FPS to 30.
    8. Go to Settings → Audio → Global Audio Devices
    9. Select Desktop Audio `Default`, all others `Disabled` 
    10. Add a new source (Display Capture) for each monitor connected. Drag and drop the sources on the canvas so they don't overlap.
    11. Go to Settings → Output and select Output Mode `Advanced`.
    12. Go to Recording and set the Recording Path to as you like (no spacings or Umlaute). Click "Generate File Name without Space" and set Recording Format to `MPEG-4 (.mp4)`.
- Download and install **Python** 3.13+:
    1. Open the installer and select "use admin privileges when installing py.exe" and "Add python.exe to PATH"
    2. Select Customize installation.
    3. Check all boxes on the next page.
    4. In "Advanced Options" check the first 5 boxes up to "precompile standard library" and then select a install location.
- Download the **Git Standalone Installer** (≥ 2.51.0) and follow its installment steps (use the recommend settings).

### 2. Clone this git repository
1. Open File explorer and go into the directory where you want the new folder with this project to be located but don't create a new folder.
2. Right-click → Open Git Bash here (if not shown, click "Show more options")
3. There run:
```bash
git clone https://github.com/LeonWeberDESY/recording_tool
```
Then go into the directory with
```bash
cd recording_tool/
```
and then execute this command:
```bash
bash install.sh
```

### 3. Update the `config.json` file using Text Editor or Visual Studio Code
1. Adjust the **OBS** websocket settings (`host`, `port`, and `password`)
2. Adjust the **OBS** recording settings (`scene_name`, `input_name`,  and `device_id`)
3. Adjust the directory and file pathes to **OBS** and the Virtual Environment **Python** executable (`pythonw.exe`)
4. Adjust the path to the **recording_tool** project folder

The path format is important, keep the double `//` for `.json` syntax.!

### 4. At this point, you can test if the setup was successfull by executing `start_obs_sipgate_controller.bat`
- Double-click the `start_obs_sipgate_controller.bat` or open a terminal in the folder and run:
```bash
.\start_obs_sipgate_controller.bat
```
- After a few seconds OBS Studio should open (you can minimize that).
- To check if the scripts are running: Open Task Manager → Details → Search for `pythonw.exe`. If there are tasks running, then the scripts are working.
- Go into Sipgate and answer a call (or go into Settings → Audio), after the recording delay this will trigger the scripts and the icon of OBS Studio in the task bar should have a red dot (recording). End the call (or leave the settings) and it should dissapear.
- Now a video file should be available at the directory you choose in OBS Studio.
- If one of these steps didn't work, check Chapter "Troubleshooting" below.

### 5. Create a boot up task using **Task Scheduler**
1. **Open Task Scheduler**
    - Press `Win + S` → type `Task Scheduler` → Enter.
2. **Create a new task**
    - Right-click Task Scheduler Library → Create Task… (not “Create Basic Task”).
3. **General tab**
    - Name: "Start OBS + Sipgate Controller" (or however you want to name it)
    - Check **Run with highest privileges** if needed.
4. **Triggers tab**
    - Click **New…** → **Begin the task: At log on** → **OK**.
    - Click **Any user**
    - Click **Delay task for 2 minutes**
    - Click **Enabled** 
5. **Actions tab**
    - Click **New…** → **Action: Start a program**
    - Program/script: `C:\Windows\System32\cmd.exe`
    - Add arguments: `/c "C:\Users\LeonWeber\Leon\recording_tool\start_obs_sipgate_controller.bat"` (change to the path of the `.bat` file in the project folder).
    - Start in: `C:\Users\LeonWeber\Leon\recording_tool` (change to your project folderpath)
6. Conditions/Settings tabs
    - Uncheck everything in the **Conditions** tab
    - Check **Allow task to be run on demand**  
    - Check **Run task as soon as possible after a scheduled start is missed** 
    - Check **If the task fails, restart every: 1 minute, Attempt o restart up to 3 times**
    - Check **Stop the task if it runs longer than 10 minutes**
    - Uncheck everything else
    - Lastly, select **Do not start a new isntance**

## Notes
- If you are working with new monitors, the first time OBS Studio opens, you have to add the desktop to one of the display captures (that is why there is a reminder pop-up).
- The microphone is dynamically created and removed during recording, ensuring the status light only triggers during active calls.
- OBS WebSocket commands rely on the input being enabled; this setup avoids leaving it always active.

## Troubleshooting
1. Ensure OBS WebSocket server is running on the configured port.
2. Verify Python packages are installed correctly (`pip list`).
3. Ensure scene and input names match exactly what’s in OBS.
4. Ensure Python is in system PATH.
5. Check that OBS is not blocked by Windows Defender or firewall.
6. If these didn't help, contact IT Expert Watson.
7. When trying to open OBS it says something like "msvcp140.dll missing", try to repair your Microsoft Visual C++ 2015-2022 Redistributable: Go to Windows Control Panel → Programs → Programs and Features → Select "Microsoft Visual C++ 2015-2022 Redistributable (x64)" → Change → Repair