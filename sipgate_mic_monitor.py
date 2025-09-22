import os
import sys
import time
import json
import logging
import subprocess
import psutil
import atexit
import threading
import queue
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL, CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED
from pycaw.pycaw import AudioUtilities, IAudioSessionManager2, IAudioSessionControl2

# -----------------------------
# Configuration
# -----------------------------
SIPGATE_PROCESS_NAME = "Sipgate.exe"
ROOT = os.path.dirname(__file__)
OBS_CONTROL_SCRIPT = os.path.join(ROOT, "obs_control.py")
PYTHON_EXE = sys.executable
LOGFILE_PATH = os.path.join(ROOT, "logs_sipgate_mic_monitor.log")
CONFIG_FILE = os.path.join(ROOT, "config.json")

with open(CONFIG_FILE, 'r') as json_file:
    cfg = json.load(json_file)
    POLL_INTERVAL = cfg.get("POLL_INTERVAL", 1)
    RECORDING_DELAY = cfg.get("RECORDING_DELAY", 3)

# -----------------------------
# Configure logging
# -----------------------------
logging.basicConfig(
    filename=LOGFILE_PATH,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

def flush_logs():
    for handler in logging.root.handlers:
        try:
            handler.flush()
        except Exception:
            pass

# -----------------------------
# Global uncaught exception hook
# -----------------------------
def log_uncaught_exceptions(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    logging.critical("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))
    flush_logs()

sys.excepthook = log_uncaught_exceptions

# -----------------------------
# STA worker thread
# -----------------------------
class STAWorker(threading.Thread):
    """
    Thread that initializes COM in STA and services mic-check requests.
    Communicates via: worker.request_q (put response_queue)
    """

    def __init__(self, sipgate_name, request_q, stop_event):
        super().__init__(daemon=True)
        self.sipgate_name = sipgate_name.lower()
        self.request_q = request_q
        self.stop_event = stop_event

    def run(self):
        logging.info("STAWorker: Starting and initializing COM (STA).")
        try:
            CoInitializeEx(COINIT_APARTMENTTHREADED)
        except Exception:
            logging.exception("STAWorker: CoInitializeEx failed (continuing to run may fail).")

        try:
            while not self.stop_event.is_set():
                try:
                    # wait for a request, or timeout to re-check stop_event
                    resp_q = self.request_q.get(timeout=0.5)
                except queue.Empty:
                    continue

                # Process request: resp_q is a queue.Queue supplied by caller
                try:
                    result = self._check_mic_active()
                    resp_q.put(("ok", result))
                except Exception as e:
                    logging.exception("STAWorker: Error while checking mic")
                    resp_q.put(("err", str(e)))
                finally:
                    # mark task done if using task_done (not strictly necessary)
                    try:
                        self.request_q.task_done()
                    except Exception:
                        pass

        finally:
            # cleanup COM
            try:
                CoUninitialize()
            except Exception:
                logging.exception("STAWorker: CoUninitialize failed.")
            logging.info("STAWorker: Exiting and uninitialized COM.")
            flush_logs()


    def _check_mic_active(self):
        """
        All COM calls must happen here (inside STA thread).
        Returns True if a SIPGATE_PROCESS_NAME session is active (state==1).
        """
        # defensive: default False
        active = False

        # enumerate devices and query sessions
        # release every COM object explicitly after use
        enumerator = None
        mic = None
        session_manager = None
        sessions = None

        try:
            enumerator = AudioUtilities.GetDeviceEnumerator()
            # role choice: eCommunications = 2 (safer if you want communications device)
            mic = enumerator.GetDefaultAudioEndpoint(1, 2)  # capture, communications
            session_manager_ptr = mic.Activate(IAudioSessionManager2._iid_, CLSCTX_ALL, None)
            session_manager = cast(session_manager_ptr, POINTER(IAudioSessionManager2))
            sessions = session_manager.GetSessionEnumerator()

            count = sessions.GetCount()
            for i in range(count):
                session = None
                session2 = None
                try:
                    session = sessions.GetSession(i)
                    session2 = session.QueryInterface(IAudioSessionControl2)
                    pid = session2.GetProcessId()
                    try:
                        proc = psutil.Process(pid)
                        if proc.name().lower() == self.sipgate_name:
                            state = session.GetState()  # 0 inactive, 1 active
                            if state == 1:
                                active = True
                                break
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        # ignore processes that disappear or are inaccessible
                        pass
                finally:
                    # release session2 and session
                    for obj in (session2, session):
                        if obj is None:
                            continue
                        try:
                            obj.Release()
                        except Exception:
                            # ignore release errors; log at debug
                            logging.debug("STAWorker: Failed to Release session object (ignored).")
                        try:
                            del obj
                        except Exception:
                            pass

            return active

        finally:
            # release top-level COM objects in reverse order
            for obj in (sessions, session_manager, mic, enumerator):
                if obj is None:
                    continue
                try:
                    obj.Release()
                except Exception:
                    logging.debug("STAWorker: Failed to Release top-level COM object (ignored).")
                try:
                    del obj
                except Exception:
                    pass

# -----------------------------
# Global STA worker and helpers
# -----------------------------
_request_q = queue.Queue()
_stop_event = threading.Event()
_sta_worker = None

def start_sta_worker():
    global _sta_worker
    if _sta_worker and _sta_worker.is_alive():
        return
    _sta_worker = STAWorker(SIPGATE_PROCESS_NAME, _request_q, _stop_event)
    _sta_worker.start()
    logging.info("Started STAWorker thread.")

def stop_sta_worker():
    _stop_event.set()
    if _sta_worker:
        _sta_worker.join(timeout=5)
    logging.info("Stopped STAWorker thread.")

atexit.register(stop_sta_worker)

def is_sipgate_mic_active_via_sta(timeout=5.0):
    """
    Request the STA worker to check mic state and wait up to `timeout` seconds for a reply.
    Returns True/False or raises RuntimeError on failure/timeout.
    """
    if _sta_worker is None or not _sta_worker.is_alive():
        # ensure worker running
        start_sta_worker()
        # small delay to let it initialize COM
        time.sleep(0.05)

    resp_q = queue.Queue(maxsize=1)
    try:
        _request_q.put(resp_q, block=True, timeout=1)
    except queue.Full:
        raise RuntimeError("STA request queue is full")

    try:
        status, payload = resp_q.get(timeout=timeout)
    except queue.Empty:
        raise RuntimeError("Timeout waiting for STA worker response")

    if status == "ok":
        return bool(payload)
    else:
        # payload contains error string
        raise RuntimeError(f"STA worker error: {payload}")

# -----------------------------
# Helper: call OBS control script
# -----------------------------
def call_obs(action):
    if action.lower() not in ("start", "stop"):
        logging.warning(f"Invalid OBS action requested: {action}")
        return
    try:
        subprocess.Popen([PYTHON_EXE, OBS_CONTROL_SCRIPT, action])
    except Exception:
        logging.exception(f"Failed to call OBS action: {action}")

# -----------------------------
# Main loop
# -----------------------------
def main():
    logging.info("Start: Monitoring Sipgate mic and controlling OBS...")
    start_sta_worker()
    recording = False
    last_mem_log = 0

    try:
        while True:
            # ask STA thread whether sipgate mic is active
            try:
                active = is_sipgate_mic_active_via_sta(timeout=3.0)
            except Exception:
                logging.exception("Error while checking Sipgate mic session")
                active = False

            try:
                if active and not recording:
                    logging.info("Call detected: Waiting for answer...")
                    stable = True
                    for _ in range(RECORDING_DELAY):
                        time.sleep(1)
                        try:
                            if not is_sipgate_mic_active_via_sta(timeout=3.0):
                                logging.info("Call not taken. Ignoring session...")
                                stable = False
                                break
                        except Exception:
                            logging.exception("Error during confirmation delay mic check")
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
                try:
                    mem = psutil.Process(os.getpid()).memory_info().rss / 1024**2
                    logging.info("Memory usage: %.2f MB", mem)
                except Exception:
                    logging.exception("Failed to get memory usage")
                last_mem_log = now
                
            flush_logs()
            time.sleep(POLL_INTERVAL)

    except KeyboardInterrupt:
        logging.error("Stopping monitoring (KeyboardInterrupt).")
    except BaseException as e:
        logging.critical(f"Fatal error in main loop: {e}", exc_info=True)
        raise
    finally:
        stop_sta_worker()
        flush_logs()

# -----------------------------
# Entry point
# -----------------------------
if __name__ == "__main__":
    main()   