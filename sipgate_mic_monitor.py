import time
import subprocess
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioSessionManager2, IAudioSessionControl2
import psutil
import sys
import os
import logging
import json

# -----------------------------
# Configuration
# -----------------------------
import json
# Open and read the JSON file
with open('config.json', 'r') as json_file:
    cfg = json.loads(json_file.read())
    POLL_INTERVAL = cfg["POLL_INTERVAL"]
    RECORDING_DELAY = cfg["RECORDING_DELAY"]

SIPGATE_PROCESS_NAME = "Sipgate.exe"
OBS_CONTROL_SCRIPT = os.path.join(os.path.dirname(__file__), "obs_control.py")
PYTHON_EXE = sys.executable  # Use the same Python interpreter
LOGFILE_PATH = os.path.join(os.path.dirname(__file__), "logs_sipgate_mic_monitor.log")

# -----------------------------
# Configure logging with timestamps
# -----------------------------
logging.basicConfig(
    filename=LOGFILE_PATH,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

# -----------------------------
# Helper: check Sipgate mic session
# -----------------------------
def is_sipgate_mic_active():
    enumerator = AudioUtilities.GetDeviceEnumerator()
    mic = enumerator.GetDefaultAudioEndpoint(1, 1)  # 1=capture, 1=communications

    session_manager = mic.Activate(IAudioSessionManager2._iid_, CLSCTX_ALL, None)
    session_manager = cast(session_manager, POINTER(IAudioSessionManager2))
    sessions = session_manager.GetSessionEnumerator()

    for i in range(sessions.GetCount()):
        session = sessions.GetSession(i)
        session2 = session.QueryInterface(IAudioSessionControl2)
        pid = session2.GetProcessId()
        try:
            proc = psutil.Process(pid)
            if proc.name().lower() == SIPGATE_PROCESS_NAME.lower():
                state = session.GetState()  # 0=inactive, 1=active
                if state == 1:
                    return True
        except Exception:
            continue
    return False

# -----------------------------
# Helper: cleanup for clearing out COM objects:
# -----------------------------
def cleanup():
    devices = AudioUtilities.GetAllDevices()
    for d in devices:
        del d

# -----------------------------
# Helper: call OBS control script
# -----------------------------
def call_obs(action):
    if action.lower() not in ("start", "stop"):
        return
    subprocess.Popen([PYTHON_EXE, OBS_CONTROL_SCRIPT, action])



# -----------------------------
# Main loop
# -----------------------------
def main():
    logging.info("Start: Monitoring Sipgate mic and controlling OBS...")
    recording = False

    try:
        while True:
            active = is_sipgate_mic_active()

            if active and not recording:
                logging.info("Call detected: Waiting for answer...")
                
                # Wait for confirmation delay while mic remains active
                stable = True
                for _ in range(RECORDING_DELAY):
                    time.sleep(1)
                    if not is_sipgate_mic_active():
                        logging.info("Call not taken. Ignoring session...")
                        stable = False
                        break

                if stable:
                    logging.info("Call answered: Forwarding OBS to RECORD")
                    call_obs("start")
                    recording = True

            elif not active and recording:
                logging.info("Call ended: Forwarding OBS to STOP")
                call_obs("stop")
                recording = False

            time.sleep(POLL_INTERVAL)

    except KeyboardInterrupt:
        logging.error("Stopping monitoring...")

    finally:
        cleanup() 
if __name__ == "__main__":
    main()
