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
from pathlib import Path
from ctypes import POINTER, cast
from comtypes import CLSCTX_ALL, CoInitialize, CoUninitialize
import tkinter as tk
from tkinter import simpledialog, messagebox
import threading

# -----------------------------
# Configuration
# -----------------------------

with open('config.json', 'r') as json_file:
    cfg = json.loads(json_file.read())
    POLL_INTERVAL = cfg["poll_interval"]
    RECORDING_DELAY = cfg["recording_delay"]
    CALL_DURATION_THRESHOLD = cfg["call_duration_threshold"] 
    RECORDING_DIR = cfg["obs_recording_path"]

SIPGATE_PROCESS_NAME = "Sipgate.exe"
OBS_CONTROL_SCRIPT = os.path.join(os.path.dirname(__file__), "obs_control.py")
PYTHON_EXE = sys.executable  # Use the same Python interpreter
LOGFILE_PATH = os.path.join(os.path.dirname(__file__), "logs_sipgate_mic_monitor.log")

# -----------------------------
# Configure logging
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
# Recording Rename Handler with Folder Organization
# -----------------------------
class RecordingRenamer:
    def __init__(self, recording_dir):
        self.recording_dir = Path(recording_dir)
        self.recording_start_time = None
        
    def mark_recording_start(self):
        """Mark the time when recording started"""
        self.recording_start_time = time.time()
        logging.info("Recording start time marked for file tracking")
    
    def get_latest_recording(self):
        """Find the most recently created recording file"""
        if not self.recording_start_time:
            logging.warning("No recording start time marked")
            return None
            
        time.sleep(2)  # Give OBS time to finalize the file
        
        # Check for common video formats
        video_extensions = ['*.mkv', '*.mp4', '*.flv', '*.mov']
        all_files = []
        
        try:
            for ext in video_extensions:
                all_files.extend(self.recording_dir.glob(ext))
            
            # Filter files created after recording started
            recent_files = [f for f in all_files if f.stat().st_mtime > self.recording_start_time]
            
            if not recent_files:
                logging.warning("No recording files found after recording start time")
                return None
            
            # Return the most recent file
            latest = max(recent_files, key=lambda x: x.stat().st_mtime)
            logging.info(f"Found latest recording: {latest.name}")
            return latest
            
        except Exception as e:
            logging.error(f"Error finding latest recording: {e}")
            return None
    
    def prompt_for_name(self, default_name=""):
        """Show a dialog to get the ticket number"""
        try:
            root = tk.Tk()
            root.withdraw()  # Hide the main window
            root.attributes('-topmost', True)  # Bring dialog to front
            
            ticket_number = simpledialog.askstring(
                "Rename Recording",
                "Enter ticket number (digits only):",
                initialvalue=default_name,
                parent=root
            )
            
            root.destroy()
            return ticket_number
        except Exception as e:
            logging.error(f"Error showing rename dialog: {e}")
            return None
    
    def get_next_recording_number(self, ticket_folder):
        """Get the next available recording number for this ticket"""
        try:
            # Find all video files in the ticket folder
            video_extensions = ['*.mkv', '*.mp4', '*.flv', '*.mov']
            existing_files = []
            
            for ext in video_extensions:
                existing_files.extend(ticket_folder.glob(ext))
            
            if not existing_files:
                return 1
            
            # Extract numbers from filenames (format: ticketnumber_XXX.ext)
            max_number = 0
            for file in existing_files:
                stem = file.stem
                # Try to extract the last part after underscore
                if '_' in stem:
                    try:
                        number_part = stem.split('_')[-1]
                        number = int(number_part)
                        max_number = max(max_number, number)
                    except ValueError:
                        continue
            
            return max_number + 1
            
        except Exception as e:
            logging.error(f"Error getting next recording number: {e}")
            return 1
    
    def organize_recording(self, old_path, ticket_number):
        """Organize recording into ticket folder with sequential numbering"""
        if not ticket_number:
            logging.info("No ticket number provided, keeping original filename")
            return old_path
        
        # Validate that ticket_number contains only digits
        if not ticket_number.isdigit():
            logging.warning(f"Ticket number '{ticket_number}' contains non-digit characters")
            try:
                messagebox.showwarning(
                    "Invalid Ticket Number",
                    "Ticket number should contain only digits.\nKeeping original filename."
                )
            except:
                pass
            return old_path
        
        try:
            # Create ticket folder if it doesn't exist
            ticket_folder = self.recording_dir / ticket_number
            ticket_folder.mkdir(exist_ok=True)
            logging.info(f"Ticket folder ready: {ticket_folder}")
            
            # Get the next recording number
            recording_number = self.get_next_recording_number(ticket_folder)
            
            # Create new filename: ticketnumber_XXX.ext
            old_extension = old_path.suffix
            new_filename = f"{ticket_number}_{recording_number:03d}{old_extension}"
            new_path = ticket_folder / new_filename
            
            # Move the file
            old_path.rename(new_path)
            logging.info(f"Recording organized: '{old_path.name}' -> '{ticket_number}/{new_filename}'")
            
            # Show success message
            try:
                messagebox.showinfo(
                    "Recording Saved",
                    f"Recording saved as:\n{ticket_number}/{new_filename}"
                )
            except:
                pass
            
            return new_path
            
        except Exception as e:
            logging.error(f"Error organizing recording: {e}")
            try:
                messagebox.showerror("Organization Error", f"Could not organize file: {e}")
            except:
                pass
            return old_path
    
    def handle_recording_rename(self):
        """Main function to handle the rename after recording stops"""
        try:
            # Find the latest recording
            latest_file = self.get_latest_recording()
            
            if not latest_file:
                logging.error("Could not find recording file to rename")
                try:
                    messagebox.showerror("Error", "Could not find the recording file")
                except:
                    pass
                return
            
            # Prompt for ticket number
            ticket_number = self.prompt_for_name()
            
            # Organize the file into ticket folder
            if ticket_number:
                self.organize_recording(latest_file, ticket_number)
            else:
                logging.info("User cancelled rename operation")
                
        except Exception as e:
            logging.error(f"Error in handle_recording_rename: {e}", exc_info=True)

# -----------------------------
# Isolated COM worker process
# -----------------------------
def com_worker_process(queue):
    """
    Separate process to handle COM operations.
    This completely isolates COM from the main process.
    Enhanced with better error handling and resource management.
    """
    # Import pycaw only in the worker process
    try:
        from pycaw.pycaw import AudioUtilities, IAudioSessionManager2, IAudioSessionControl2
    except ImportError as e:
        print(f"COM worker failed to import pycaw: {e}")
        return
    
    # Initialize COM for this process
    try:
        pythoncom.CoInitialize()
    except Exception as e:
        print(f"COM worker failed to initialize: {e}")
        return
    
    enumerator = None
    mic = None
    session_manager = None
    
    try:
        # Pre-initialize COM objects to reduce per-check overhead
        enumerator = AudioUtilities.GetDeviceEnumerator()
        mic = enumerator.GetDefaultAudioEndpoint(1, 1)  # 1=capture, 1=communications
        session_manager = mic.Activate(IAudioSessionManager2._iid_, CLSCTX_ALL, None)
        session_manager = cast(session_manager, POINTER(IAudioSessionManager2))
        
        check_count = 0
        max_checks = 400  # Restart worker after 400 checks to prevent memory issues
        
        while check_count < max_checks:
            try:
                # Get command from queue
                command = queue.get(timeout=2)
                
                if command == "CHECK":
                    result = False
                    check_count += 1
                    
                    try:
                        # Get fresh session list each time
                        sessions = session_manager.GetSessionEnumerator()
                        
                        for i in range(sessions.GetCount()):
                            try:
                                session = sessions.GetSession(i)
                                session2 = session.QueryInterface(IAudioSessionControl2)
                                pid = session2.GetProcessId()
                                
                                # Quick process name check
                                try:
                                    proc = psutil.Process(pid)
                                    if proc.name().lower() == SIPGATE_PROCESS_NAME.lower():
                                        state = session.GetState()  # 0=inactive, 1=active
                                        if state == 1:
                                            result = True
                                            break
                                except (psutil.NoSuchProcess, psutil.AccessDenied):
                                    continue
                                    
                            except Exception:
                                # Skip problematic sessions
                                continue
                        
                        queue.put(result)
                        
                        # Periodic cleanup to prevent memory accumulation
                        if check_count % 50 == 0:
                            import gc
                            gc.collect()
                        
                    except Exception as e:
                        queue.put(False)
                        print(f"COM worker check error: {e}")
                        
                elif command == "EXIT":
                    print("COM worker received EXIT command")
                    break
                    
            except Exception:
                # Timeout or other queue error - continue waiting
                continue
        
        print(f"COM worker exiting after {check_count} checks (planned restart)")
                
    except Exception as e:
        print(f"COM worker fatal error: {e}")
        
    finally:
        # Cleanup COM objects
        try:
            if session_manager:
                del session_manager
            if mic:
                del mic
            if enumerator:
                del enumerator
        except:
            pass
        
        # Cleanup COM
        try:
            pythoncom.CoUninitialize()
        except:
            pass

# -----------------------------
# Process-based mic checker with lifecycle management
# -----------------------------
class ProcessSafeMicChecker:
    def __init__(self):
        self.process = None
        self.queue = None
        self.response_queue = None
        self.worker_start_time = None
        self.check_count = 0
        self.max_worker_lifetime = 600  # 10 minutes before forced restart
        self.max_checks_per_worker = 500  # Restart after 500 checks
        self.consecutive_errors = 0
        self.max_consecutive_errors = 3
        self.start_worker()
    
    def start_worker(self):
        """Start the COM worker process."""
        self.cleanup()  # Ensure clean state
        
        self.queue = multiprocessing.Queue()
        self.response_queue = multiprocessing.Queue()
        
        self.process = multiprocessing.Process(target=com_worker_process, args=(self.queue,))
        self.process.daemon = True
        self.process.start()
        self.worker_start_time = time.time()
        self.check_count = 0
        self.consecutive_errors = 0
        logging.info("Started COM worker process")
    
    def should_restart_worker(self):
        """Check if worker should be restarted for maintenance."""
        if not self.process or not self.process.is_alive():
            return True
        
        current_time = time.time()
        
        # Restart if worker has been running too long
        if self.worker_start_time and (current_time - self.worker_start_time) > self.max_worker_lifetime:
            logging.info(f"Worker restart due to age: {current_time - self.worker_start_time:.0f}s")
            return True
        
        # Restart if too many checks have been performed
        if self.check_count > self.max_checks_per_worker:
            logging.info(f"Worker restart due to check count: {self.check_count}")
            return True
        
        # Restart if too many consecutive errors
        if self.consecutive_errors >= self.max_consecutive_errors:
            logging.info(f"Worker restart due to consecutive errors: {self.consecutive_errors}")
            return True
        
        return False
    
    def check_mic(self):
        """Check if Sipgate mic is active using the worker process."""
        # Proactive worker maintenance
        if self.should_restart_worker():
            logging.info("Performing preventive worker restart...")
            self.start_worker()
            time.sleep(1)  # Give it time to initialize
        
        try:
            # Clear any old responses with timeout
            clear_attempts = 0
            while not self.queue.empty() and clear_attempts < 10:
                try:
                    self.queue.get_nowait()
                    clear_attempts += 1
                except:
                    break
            
            # Send check command
            self.queue.put("CHECK")
            self.check_count += 1
            
            # Wait for response (with timeout)
            result = self.queue.get(timeout=3)
            self.consecutive_errors = 0  # Reset on success
            return result
            
        except Exception as e:
            self.consecutive_errors += 1
            logging.error(f"Error checking mic (consecutive: {self.consecutive_errors}): {e}")
            
            # Don't restart immediately on first error, give it a chance
            if self.consecutive_errors >= 2:
                logging.warning("Multiple consecutive errors, restarting worker...")
                self.start_worker()
                time.sleep(1)
            
            return False
    
    def cleanup(self):
        """Clean up the worker process."""
        if self.process and self.process.is_alive():
            try:
                # Try graceful shutdown first
                if hasattr(self, 'queue') and self.queue:
                    try:
                        self.queue.put("EXIT", timeout=1)
                    except:
                        pass
                
                # Wait for graceful exit
                self.process.join(timeout=3)
                
                # Force termination if needed
                if self.process.is_alive():
                    logging.warning("Force terminating unresponsive worker...")
                    self.process.terminate()
                    self.process.join(timeout=3)
                    
                    # Kill if still alive
                    if self.process.is_alive():
                        logging.warning("Force killing worker process...")
                        self.process.kill()
                        self.process.join(timeout=2)
                        
            except Exception as e:
                logging.error(f"Error during worker cleanup: {e}")
            
            finally:
                self.process = None
                
        # Clean up queues
        if hasattr(self, 'queue') and self.queue:
            try:
                self.queue.close()
            except:
                pass
        if hasattr(self, 'response_queue') and self.response_queue:
            try:
                self.response_queue.close()
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
def call_obs(action, renamer=None):
    if action.lower() not in ("start", "stop"):
        return
    try:
        subprocess.Popen([PYTHON_EXE, OBS_CONTROL_SCRIPT, action])
        logging.info(f"Called OBS control with action: {action}")
        
        # Mark recording start time for file tracking
        if action.lower() == "start" and renamer:
            renamer.mark_recording_start()
            
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
    
    # Create the recording renamer
    renamer = RecordingRenamer(RECORDING_DIR)
    
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
                            call_obs("start", renamer)
                            recording = True
                
                else:  # No microphone activity
                    if call_start_time is not None:
                        call_duration = current_time - call_start_time
                        
                        if call_duration < CALL_DURATION_THRESHOLD:
                            # Audio was too brief, ignore it
                            pass
                        else:
                            # This was a real call that has now ended
                            if recording:
                                logging.info("Call ended: Stopping OBS recording")
                                call_obs("stop")
                                recording = False
                                
                                # Launch rename dialog in a separate thread to avoid blocking
                                time.sleep(2)  # Give OBS time to finalize the file
                                logging.info("Launching rename dialog...")
                                rename_thread = threading.Thread(target=renamer.handle_recording_rename)
                                rename_thread.daemon = True
                                rename_thread.start()
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