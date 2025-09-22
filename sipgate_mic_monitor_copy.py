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
# Global uncaught exception hook
# -----------------------------
def log_uncaught_exceptions(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    logging.critical("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))

sys.excepthook = log_uncaught_exceptions

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
    last_mem_log = 0

    try:
        while True:
            try:
                active = is_sipgate_mic_active()
            except Exception:
                logging.exception("Error while checking Sipgate mic session")
                active = False

            try:
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

            except Exception:
                logging.exception("Error while handling call state logic")

            # Periodic memory usage logging (every 1 minute)
            now = time.time()
            if now - last_mem_log >= 60:
                mem = psutil.Process(os.getpid()).memory_info().rss / 1024**2
                logging.info(f"Memory usage: {mem:.2f} MB")
                last_mem_log = now

            time.sleep(POLL_INTERVAL)

    except KeyboardInterrupt:
        logging.error("Stopping monitoring (KeyboardInterrupt).")

    finally:
        cleanup()

# -----------------------------
# Entry point with top-level error logging
# -----------------------------
if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        logging.critical(f"SystemExit: {e}")
        raise
    except OSError as e:
        logging.critical(f"OSError caused shutdown: {e}", exc_info=True)
        raise
    except Exception as e:
        logging.exception("Fatal error in main loop")
        raise
