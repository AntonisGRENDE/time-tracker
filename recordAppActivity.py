import os
import signal
import time
from threading import Thread

import openpyxl
import psutil
import win32gui
import win32process
from PIL import Image, ImageDraw
from openpyxl.utils import get_column_letter
from pystray import Icon, MenuItem as item

# Path to the log file
log_file_path = "app_usage_log.xlsx"

# Initialize usage records
usage_records = {}

sleepingTime = 2

# Global flag to stop threads
stop_threads = False

def signal_handler(sig, frame):
    global stop_threads
    print("\n[INFO] Stopping application...")
    stop_threads = True
    icon.stop()  # Stop the system tray icon
    # Allow threads to exit naturally without calling sys.exit()


# Register the signal handler for Ctrl + C
signal.signal(signal.SIGINT, signal_handler)

# Function to format duration as HH:MM:SS
def format_duration(seconds):
    hours, seconds = divmod(seconds, 3600)
    minutes, seconds = divmod(seconds, 60)
    return f"{int(hours):02}:{int(minutes):02}:{int(seconds):02}"


# Function to log app usage
def log_usage(app_name, app_title, start_time, end_time, duration):
    global usage_records
    key = (app_name, app_title)  # Use a tuple as the key
    if key in usage_records:
        # Update the duration for existing records
        usage_records[key] += duration
    else:
        # Add a new record
        usage_records[key] = duration
    print(f"Logged: {app_name}, {app_title}, Total Duration: {format_duration(usage_records[key])}")



def write_to_file():
    global usage_records, stop_threads
    while not stop_threads:
        if usage_records:
            try:
                consolidated_records = {}

                # Check if the file exists
                if os.path.exists(log_file_path):
                    # Load the workbook and read existing data
                    wb = openpyxl.load_workbook(log_file_path)
                    ws = wb.active

                    # Populate consolidated_records with data from the file
                    for row in ws.iter_rows(min_row=2, values_only=True):  # Skip the header row
                        app_name, app_title, duration = row
                        if app_name and app_title:
                            # Convert duration from HH:MM:SS to seconds
                            hours, minutes, seconds = map(int, duration.split(':'))
                            duration_in_seconds = hours * 3600 + minutes * 60 + seconds
                            key = (app_name, app_title)
                            if key in consolidated_records:
                                consolidated_records[key] += duration_in_seconds
                            else:
                                consolidated_records[key] = duration_in_seconds
                else:
                    # Create a new file if it doesn't exist
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "App Usage Log"
                    ws.append(["App Name", "App Title", "Duration"])  # Add headers

                # Merge in-memory records with consolidated records
                for (app_name, app_title), duration in usage_records.items():
                    key = (app_name, app_title)
                    if key in consolidated_records:
                        consolidated_records[key] += duration
                    else:
                        consolidated_records[key] = duration

                # Write consolidated records back to the file
                ws.delete_rows(2, ws.max_row)  # Clear existing data but keep the headers
                for (app_name, app_title), total_duration in consolidated_records.items():
                    ws.append([app_name, app_title, format_duration(total_duration)])

                # Autofit columns
                for col in ws.columns:
                    max_length = 0
                    col_letter = get_column_letter(col[0].column)  # Get column letter
                    for cell in col:
                        try:
                            if cell.value:
                                max_length = max(max_length, len(str(cell.value)))
                        except:
                            pass
                    adjusted_width = max_length + 2  # Add some padding
                    ws.column_dimensions[col_letter].width = adjusted_width

                # Save the updated workbook
                wb.save(log_file_path)
                print(f"[INFO] Records written to {log_file_path}")
            except PermissionError:
                print(f"[Warning] Unable to write to {log_file_path}. Retrying...")
        time.sleep(sleepingTime)




# Function to track app usage
def track_app_usage():
    last_app = None
    last_title = None
    start_time = None

    while not stop_threads:
        try:
            # Get the handle of the foreground window
            hwnd = win32gui.GetForegroundWindow()
            if hwnd:
                # Get process ID from the window handle
                _, process_id = win32process.GetWindowThreadProcessId(hwnd)
                process = psutil.Process(process_id)
                current_app = process.name()
                current_title = win32gui.GetWindowText(hwnd)
            else:
                current_app, current_title = "Unknown", "No Active Window"

            # Detect app switch
            if current_app != last_app or current_title != last_title:
                if last_app and start_time:
                    # Calculate usage duration
                    end_time = time.time()
                    duration = end_time - start_time
                    log_usage(last_app, last_title, time.ctime(start_time), time.ctime(end_time), duration)

                # Update last app details
                last_app = current_app
                last_title = current_title
                start_time = time.time()

            # Sleep for a longer duration to save power (adjust as needed)
            time.sleep(sleepingTime)
        except Exception as e:
            print(f"[Error] {e}")
            time.sleep(sleepingTime)  # Retry after 5 seconds

# Create a system tray icon
def create_image(width, height, color1, color2):
    image = Image.new('RGB', (width, height), color1)
    dc = ImageDraw.Draw(image)
    dc.rectangle(
        (width // 2, 0, width, height // 2),
        fill=color2)
    dc.rectangle(
        (0, height // 2, width // 2, height),
        fill=color2)
    return image

def quit_application(icon, item):
    icon.stop()

def setup(icon):
    icon.visible = True

# Start tracking in a separate thread
tracking_thread = Thread(target=track_app_usage, daemon=True)
tracking_thread.start()

# Start writing to file in a separate thread
write_thread = Thread(target=write_to_file, daemon=True)
write_thread.start()

# Create and run system tray icon
menu = (
    item('Quit', quit_application),
)
icon_image = create_image(64, 64, 'black', 'red')
icon = Icon("App Tracker", icon_image, "App Usage Tracker", menu)

# Run the system tray icon
try:
    icon.run(setup)
except Exception as e:
    print(f"[Error] {e}")