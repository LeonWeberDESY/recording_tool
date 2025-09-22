import time
import subprocess
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL, CoInitialize, CoUninitialize
from pycaw.pycaw import AudioUtilities, IAudioSessionManager2, IAudioSessionControl2
import psutil
import sys
import os
import logging
import json
import atexit
import gc
import weakref

# -----------------------------
# Configuration
# -----------------------------

with open('config.json', 'r') as json_file:
    cfg = json.loads(json_file.read())
    POLL_INTERVAL = cfg["POLL_INTERVAL"]
    RECORDING_DELAY = cfg["RECORDING_DELAY"]

SIPGATE_PROCESS_NAME = "Sipgate.exe"
OBS_CONTROL_SCRIPT = os.path.join(os.path.dirname(__file__), "obs_control.py")
PYTHON_EXE = sys.executable
LOGFILE_PATH = os.path.join(os.path.dirname(__file__), "logs_sipgate_mic_monitor.log")

# -----------------------------
# COM object tracking for proper cleanup
# -----------------------------
_com_objects = weakref.WeakSet()

def track_com_object(obj):
    """Track COM objects for proper cleanup"""
    if obj is not None:
        _com_objects.add(obj)
    return obj

def cleanup_com_objects():
    """Safely cleanup all tracked COM objects"""
    objects_to_cleanup = list(_com_objects)
    for obj in objects_to_cleanup:
        try:
            if hasattr(obj, 'Release'):
                obj.Release()
        except:
            pass  # Ignore cleanup errors
    _com_objects.clear()
    gc.collect()

# -----------------------------
# Configure logging with timestamps
# -----------------------------
logging.basicConfig(
    filename=LOGFILE_PATH,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

# Also log to console for debugging
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
console_handler.setFormatter(formatter)
logging.getLogger().addHandler(console_handler)

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
# COM initialization and cleanup
# -----------------------------
def initialize_com():
    """Initialize COM for the current thread"""
    try:
        CoInitialize()
        logging.info("COM initialized successfully")
        return True
    except Exception as e:
        logging.error(f"Failed to initialize COM: {e}")
        return False

def cleanup_com():
    """Cleanup COM objects and uninitialize COM"""
    try:
        cleanup_com_objects()
        CoUninitialize()
        logging.info("COM cleanup completed")
    except Exception as e:
        logging.warning(f"COM cleanup warning: {e}")

# Register cleanup function
atexit.register(cleanup_com)

# -----------------------------
# Helper: check Sipgate mic session with improved error handling
# -----------------------------
def is_sipgate_mic_active():
    enumerator = None
    mic = None
    session_manager = None
    sessions = None
    
    try:
        enumerator = track_com_object(AudioUtilities.GetDeviceEnumerator())
        if not enumerator:
            return False
            
        mic = track_com_object(enumerator.GetDefaultAudioEndpoint(1, 1))  # 1=capture, 1=communications
        if not mic:
            return False

        session_manager = track_com_object(mic.Activate(IAudioSessionManager2._iid_, CLSCTX_ALL, None))
        session_manager = cast(session_manager, POINTER(IAudioSessionManager2))
        
        sessions = track_com_object(session_manager.GetSessionEnumerator())
        if not sessions:
            return False

        session_count = sessions.GetCount()
        for i in range(session_count):
            session = None
            session2 = None
            try:
                session = track_com_object(sessions.GetSession(i))
                if not session:
                    continue
                    
                session2 = track_com_object(session.QueryInterface(IAudioSessionControl2))
                if not session2:
                    continue
                    
                pid = session2.GetProcessId()
                if pid == 0:  # System process
                    continue
                    
                proc = psutil.Process(pid)
                if proc.name().lower() == SIPGATE_PROCESS_NAME.lower():
                    state = session.GetState()  # 0=inactive, 1=active
                    if state == 1:
                        return True
                        
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                # Process no longer exists or can't be accessed
                continue
            except Exception as e:
                logging.debug(f"Error checking session {i}: {e}")
                continue
            finally:
                # Clean up session objects immediately
                if session2:
                    try:
                        session2.Release()
                    except:
                        pass
                if session:
                    try:
                        session.Release()
                    except:
                        pass
                        
    except Exception as e:
        logging.error(f"Error in is_sipgate_mic_active: {e}")
        return False
    finally:
        # Clean up main objects
        for obj in [sessions, session_manager, mic, enumerator]:
            if obj:
                try:
                    obj.Release()
                except:
                    pass
        
        # Force garbage collection to help with COM cleanup
        gc.collect()
    
    return False

# -----------------------------
# Helper: call OBS control script with error handling
# -----------------------------
def call_obs(action):
    if action.lower() not in ("start", "stop"):
        logging.warning(f"Invalid OBS action: {action}")
        return
    
    try:
        # Use subprocess.run with timeout for better error handling
        result = subprocess.run(
            [PYTHON_EXE, OBS_CONTROL_SCRIPT, action],
            timeout=30,  # 30 second timeout
            capture_output=True,
            text=True
        )
        
        if result.returncode != 0:
            logging.error(f"OBS control script failed with return code {result.returncode}")
            if result.stderr:
                logging.error(f"OBS script stderr: {result.stderr}")
        else:
            logging.info(f"OBS {action} command executed successfully")
            
    except subprocess.TimeoutExpired:
        logging.error(f"OBS control script timed out for action: {action}")
    except Exception as e:
        logging.error(f"Error calling OBS script: {e}")

# -----------------------------
# Main loop with improved error handling
# -----------------------------
def main():
    logging.info("Starting Sipgate mic monitor...")
    
    if not initialize_com():
        logging.critical("Failed to initialize COM. Exiting.")
        return 1
    
    recording = False
    last_mem_log = 0
    consecutive_errors = 0
    max_consecutive_errors = 5

    try:
        while True:
            try:
                active = is_sipgate_mic_active()
                consecutive_errors = 0  # Reset error counter on success
                
            except Exception as e:
                consecutive_errors += 1
                logging.exception(f"Error while checking Sipgate mic session (attempt {consecutive_errors})")
                
                if consecutive_errors >= max_consecutive_errors:
                    logging.critical(f"Too many consecutive errors ({consecutive_errors}). Restarting COM...")
                    cleanup_com()
                    time.sleep(5)  # Wait before reinitializing
                    if not initialize_com():
                        logging.critical("Failed to reinitialize COM. Exiting.")
                        return 1
                    consecutive_errors = 0
                
                active = False
                time.sleep(POLL_INTERVAL * 2)  # Wait longer after error
                continue

            try:
                if active and not recording:
                    logging.info("Call detected: Waiting for answer...")

                    # Wait for confirmation delay while mic remains active
                    stable = True
                    for i in range(RECORDING_DELAY):
                        time.sleep(1)
                        try:
                            if not is_sipgate_mic_active():
                                logging.info(f"Call not taken after {i+1}s. Ignoring session...")
                                stable = False
                                break
                        except Exception as e:
                            logging.warning(f"Error during stability check: {e}")
                            stable = False
                            break

                    if stable:
                        logging.info("Call answered: Starting OBS recording")
                        call_obs("start")
                        recording = True

                elif not active and recording:
                    logging.info("Call ended: Stopping OBS recording")
                    call_obs("stop")
                    recording = False

            except Exception as e:
                logging.exception("Error while handling call state logic")

            # Periodic memory usage logging (every 1 minute)
            now = time.time()
            if now - last_mem_log >= 60:
                try:
                    current_process = psutil.Process(os.getpid())
                    mem = current_process.memory_info().rss / 1024**2
                    cpu = current_process.cpu_percent()
                    logging.info(f"Memory usage: {mem:.2f} MB, CPU: {cpu:.1f}%")
                    last_mem_log = now
                except Exception as e:
                    logging.warning(f"Error getting process info: {e}")

            time.sleep(POLL_INTERVAL)

    except KeyboardInterrupt:
        logging.info("Stopping monitoring (KeyboardInterrupt)")
        if recording:
            logging.info("Stopping final recording...")
            call_obs("stop")
            
    except Exception as e:
        logging.critical(f"Fatal error in main loop: {e}", exc_info=True)
        return 1

    finally:
        cleanup_com()

    return 0

# -----------------------------
# Entry point with improved error handling
# -----------------------------
if __name__ == "__main__":
    try:
        exit_code = main()
        sys.exit(exit_code)
    except SystemExit:
        raise
    except Exception as e:
        logging.exception("Fatal error at top level")
        cleanup_com()
        sys.exit(1)