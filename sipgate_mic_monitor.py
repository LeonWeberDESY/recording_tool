import time
import subprocess
import psutil
import sys
import os
import logging
import json
import gc
import multiprocessing
from multiprocessing import Queue
import pythoncom
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL, CoInitialize, CoUninitialize

# -----------------------------
# Configuration
# -----------------------------

with open('config.json', 'r') as json_file:
    cfg = json.loads(json_file.read())
    POLL_INTERVAL = cfg["POLL_INTERVAL"]
    RECORDING_DELAY = cfg["RECORDING_DELAY"]
    CALL_DURATION_THRESHOLD = cfg["CALL_DURATION_THRESHOLD"] 

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

# Also log to console for debugging
console = logging.StreamHandler()
console.setLevel(logging.INFO)
logging.getLogger('').addHandler(console)

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
# Isolated COM worker process
# -----------------------------
def com_worker_process(queue):
    """
    Separate process to handle COM operations.
    This completely isolates COM from the main process.
    """
    # Import pycaw only in the worker process
    from pycaw.pycaw import AudioUtilities, IAudioSessionManager2, IAudioSessionControl2
    
    # Initialize COM for this process
    pythoncom.CoInitialize()
    
    try:
        while True:
            try:
                # Get command from queue
                command = queue.get(timeout=1)
                
                if command == "CHECK":
                    result = False
                    
                    try:
                        # Perform the mic check
                        enumerator = AudioUtilities.GetDeviceEnumerator()
                        mic = enumerator.GetDefaultAudioEndpoint(1, 1)  # 1=capture, 1=communications
                        
                        session_manager = mic.Activate(IAudioSessionManager2._iid_, CLSCTX_ALL, None)
                        session_manager = cast(session_manager, POINTER(IAudioSessionManager2))
                        sessions = session_manager.GetSessionEnumerator()
                        
                        for i in range(sessions.GetCount()):
                            try:
                                session = sessions.GetSession(i)
                                session2 = session.QueryInterface(IAudioSessionControl2)
                                pid = session2.GetProcessId()
                                
                                proc = psutil.Process(pid)
                                if proc.name().lower() == SIPGATE_PROCESS_NAME.lower():
                                    state = session.GetState()  # 0=inactive, 1=active
                                    if state == 1:
                                        result = True
                                        break
                            except:
                                continue
                        
                        queue.put(result)
                        
                    except Exception as e:
                        queue.put(False)
                        print(f"COM worker error: {e}")
                        
                elif command == "EXIT":
                    break
                    
            except:
                # Timeout - continue waiting
                continue
                
    finally:
        pythoncom.CoUninitialize()

# -----------------------------
# Process-based mic checker
# -----------------------------
class ProcessSafeMicChecker:
    def __init__(self):
        self.process = None
        self.queue = None
        self.response_queue = None
        self.start_worker()
    
    def start_worker(self):
        """Start the COM worker process."""
        self.queue = multiprocessing.Queue()
        self.response_queue = multiprocessing.Queue()
        
        # Create wrapper to handle both queues
        def worker_wrapper():
            com_worker_process(self.queue)
        
        self.process = multiprocessing.Process(target=com_worker_process, args=(self.queue,))
        self.process.daemon = True
        self.process.start()
        logging.info("Started COM worker process")
    
    def check_mic(self):
        """Check if Sipgate mic is active using the worker process."""
        if not self.process or not self.process.is_alive():
            logging.warning("Worker process dead, restarting...")
            self.cleanup()
            self.start_worker()
            time.sleep(1)  # Give it time to initialize
        
        try:
            # Clear any old responses
            while not self.queue.empty():
                try:
                    self.queue.get_nowait()
                except:
                    break
            
            # Send check command
            self.queue.put("CHECK")
            
            # Wait for response (with timeout)
            result = self.queue.get(timeout=3)
            return result
            
        except Exception as e:
            logging.error(f"Error checking mic: {e}")
            # Restart worker on error
            self.cleanup()
            self.start_worker()
            return False
    
    def cleanup(self):
        """Clean up the worker process."""
        if self.process and self.process.is_alive():
            try:
                self.queue.put("EXIT")
                self.process.join(timeout=2)
                if self.process.is_alive():
                    self.process.terminate()
                    self.process.join(timeout=2)
                    if self.process.is_alive():
                        self.process.kill()
            except:
                pass

# -----------------------------
# Alternative: WMI-based checker (fallback)
# -----------------------------
def check_sipgate_mic_wmi():
    """
    Alternative method using WMI to check audio sessions.
    This is a fallback if COM continues to fail.
    """
    try:
        import wmi
        c = wmi.WMI()
        
        # Check if Sipgate process exists and has audio
        for process in c.Win32_Process(Name=SIPGATE_PROCESS_NAME):
            # If process exists, we assume it might be in a call
            # This is less accurate but more stable
            return True
        return False
    except:
        return False

# -----------------------------
# Helper: call OBS control script
# -----------------------------
def call_obs(action):
    if action.lower() not in ("start", "stop"):
        return
    try:
        subprocess.Popen([PYTHON_EXE, OBS_CONTROL_SCRIPT, action])
        logging.info(f"Called OBS control with action: {action}")
    except Exception as e:
        logging.error(f"Failed to call OBS control: {e}")

# -----------------------------
# Main loop with process isolation and duration threshold
# -----------------------------
def main():
    logging.info("="*50)
    logging.info("Starting Sipgate mic monitor (Process-Isolated Version)")
    logging.info(f"Python version: {sys.version}")
    logging.info(f"Poll interval: {POLL_INTERVAL}s, Recording delay: {RECORDING_DELAY}s")
    logging.info(f"Call duration threshold: {CALL_DURATION_THRESHOLD}s")
    logging.info("="*50)
    
    recording = False
    last_mem_log = 0
    error_count = 0
    max_consecutive_errors = 10
    use_fallback = False
    
    # Duration threshold variables
    call_start_time = None
    call_detection_logged = False
    
    # Create the process-safe mic checker
    mic_checker = ProcessSafeMicChecker()
    
    try:
        while True:
            current_time = time.time()
            
            try:
                # Check mic status
                if not use_fallback:
                    try:
                        active = mic_checker.check_mic()
                        error_count = 0  # Reset on success
                    except Exception as e:
                        error_count += 1
                        logging.error(f"Mic check failed ({error_count}/{max_consecutive_errors}): {e}")
                        
                        if error_count >= max_consecutive_errors:
                            logging.warning("Too many COM errors, switching to fallback method")
                            use_fallback = True
                            active = check_sipgate_mic_wmi()
                        else:
                            active = False
                else:
                    # Use fallback WMI method
                    active = check_sipgate_mic_wmi()
                
            except Exception as e:
                logging.exception(f"Critical error in mic check: {e}")
                active = False

            try:
                # Handle audio activity with duration threshold
                if active:
                    if call_start_time is None:
                        # First detection of audio activity
                        call_start_time = current_time
                        call_detection_logged = False
                        
                    
                    # Check if audio has been active long enough to be considered a call
                    call_duration = current_time - call_start_time
                    
                    if call_duration >= CALL_DURATION_THRESHOLD:
                        # This is now considered a real call
                        if not call_detection_logged:
                            logging.info("Audio activity detected: Checking duration...")
                            call_detection_logged = True
                        
                        # Check if we should start recording (after recording delay)
                        if not recording and call_duration >= RECORDING_DELAY:
                            logging.info("Call answered: Starting OBS recording")
                            call_obs("start")
                            recording = True
                
                else:  # No microphone activity
                    if call_start_time is not None:
                        call_duration = current_time - call_start_time
                        
                        if call_duration < CALL_DURATION_THRESHOLD:
                            # Audio was too brief, ignore it
                            #logging.info(f"Brief audio activity ignored ({call_duration:.1f}s < {CALL_DURATION_THRESHOLD}s threshold)")
                            pass
                        else:
                            # This was a real call that has now ended
                            if recording:
                                logging.info("Call ended: Stopping OBS recording")
                                call_obs("stop")
                                recording = False
                            else:
                                logging.info(f"Call session ended without recording (duration: {call_duration:.1f}s)")
                        
                        # Reset call tracking variables
                        call_start_time = None
                        call_detection_logged = False

            except Exception as e:
                logging.exception(f"Error in call state logic: {e}")

            # Periodic status logging
            if current_time - last_mem_log >= 60:
                try:
                    process = psutil.Process(os.getpid())
                    mem_mb = process.memory_info().rss / 1024**2
                    cpu_percent = process.cpu_percent(interval=0.1)
                    method = "Fallback/WMI" if use_fallback else "COM/Process"
                    logging.info(f"Status - Memory: {mem_mb:.2f} MB, CPU: {cpu_percent:.1f}%, "
                               f"Recording: {recording}, Method: {method}")
                except Exception as e:
                    logging.error(f"Could not log stats: {e}")
                last_mem_log = current_time

            time.sleep(POLL_INTERVAL)

    except KeyboardInterrupt:
        logging.info("Stopping monitoring (KeyboardInterrupt)")

    finally:
        logging.info("Performing cleanup...")
        mic_checker.cleanup()
        logging.info("Sipgate mic monitor stopped")

# -----------------------------
# Robust entry point
# -----------------------------
if __name__ == "__main__":
    # Set multiprocessing start method for Windows
    if sys.platform == "win32":
        multiprocessing.set_start_method('spawn', force=True)
    
    max_restarts = 5
    restart_count = 0
    
    while restart_count < max_restarts:
        try:
            main()
            break
            
        except SystemExit:
            logging.info("Normal exit")
            break
            
        except Exception as e:
            restart_count += 1
            logging.critical(f"Fatal error (restart {restart_count}/{max_restarts}): {e}", exc_info=True)
            
            if restart_count < max_restarts:
                wait_time = min(30 * restart_count, 300)
                logging.info(f"Restarting in {wait_time} seconds...")
                time.sleep(wait_time)
            else:
                logging.critical("Maximum restarts reached. Exiting.")
                
                # Create a marker file to indicate repeated crashes
                try:
                    with open("CRASH_MARKER.txt", "w") as f:
                        f.write(f"Crashed {max_restarts} times. Last error: {str(e)}\n")
                        f.write(f"Time: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                except:
                    pass
                    
                sys.exit(1)