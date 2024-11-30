import time
import csv
import os
import psutil
import pygetwindow as gw
from pystray import Icon, MenuItem as item
from PIL import Image, ImageDraw
from threading import Thread
import win32gui
import win32process
import signal
import sys
import msvcrt

# Path to the log file
log_file_path = "app_usage_log.csv"

# Initialize usage records
usage_records = []

sleepingTime = 2

# Global flag to stop threads
stop_threads = False

# Signal handler for Ctrl + C
def signal_handler(sig, frame):
    global stop_threads
    print("\n[INFO] Stopping application...")
    stop_threads = True
    icon.stop()  # Stop the system tray icon
    sys.exit(0)

# Register the signal handler for Ctrl + C
signal.signal(signal.SIGINT, signal_handler)

# Function to format duration as HH:MM:SS
def format_duration(seconds):
    hours, seconds = divmod(seconds, 3600)
    minutes, seconds = divmod(seconds, 60)
    return f"{int(hours):02}:{int(minutes):02}:{int(seconds):02}"

# Function to write usage data to the file periodically
def write_to_file():
    global usage_records, stop_threads
    while not stop_threads:  # Check stop_threads flag
        if usage_records:
            try:
                with open(log_file_path, mode='a', newline='') as file:
                    # Enable shared access for the file
                    msvcrt.locking(file.fileno(), msvcrt.LK_NBLCK, 1)
                    writer = csv.writer(file)
                    writer.writerows(usage_records)
                    usage_records = []  # Clear records after writing
                    msvcrt.locking(file.fileno(), msvcrt.LK_UNLCK, 1)
            except PermissionError:
                print(f"[Warning] Unable to write to {log_file_path}. Retrying...")
            except OSError as e:
                print(f"[Error] OS Error: {e}")
        time.sleep(sleepingTime)  # Write every 2 seconds

# Function to log app usage
def log_usage(app_name, app_title, start_time, end_time, duration):
    formatted_duration = format_duration(duration)
    usage_records.append([app_name, app_title, formatted_duration, start_time, end_time])
    print(f"Logged: {app_name}, {app_title}")


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
