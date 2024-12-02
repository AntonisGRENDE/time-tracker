import os
import signal
import sys
import time
from threading import Thread

import openpyxl
import psutil
import win32gui
import win32process
from PIL import Image, ImageDraw
from openpyxl.utils import get_column_letter
from pystray import Icon, MenuItem as item
from datetime import datetime
from threading import Lock


#timestamp = datetime.now().strftime("%M.%Y_%H:%m")
timestampFile = datetime.now().strftime("%H-%M_%m.%Y")  # Month.Year_Hour-Minute
# Path to the log file
log_file_path = f"app_usage_log_{timestampFile}.xlsx"
print (log_file_path)
print ("application  is starting at " + str(datetime.now().strftime("%H:%M")))


print("Current open window:" + win32gui.GetWindowText(win32gui.GetForegroundWindow()))


# Initialize usage records and locks
usage_records = {}
usage_lock = Lock()
icon = None
sleepingTime = 2
sleepingTimeWriteToFile = 4

# Global flag to stop threads
stop_threads = False

def signal_handler(sig, frame):
    global stop_threads, icon
    print("\n[INFO] Stopping application and finalizing data...")
    stop_threads = True

    # Stop tracking and writing threads
    tracking_thread.join(timeout=2)
    write_thread.join(timeout=2)

    # Aggregate data
    aggregate_detailed_usage()

    # Stop system tray icon
    if icon:
        icon.stop()

    os.startfile(log_file_path)


def create_workbook():
    """Ensure a clean workbook is created with a single Session_Log sheet."""
    wb = openpyxl.Workbook()
    # Remove default sheet
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    # Create Session_Log sheet
    ws = wb.create_sheet(title="Session_Log")
    ws.append(["App Name", "App Title", "Duration", "Timestamp"])

    wb.save(log_file_path)
    return wb

# Function to format duration as HH:MM:SS
def format_duration(seconds):
    hours, seconds = divmod(seconds, 3600)
    minutes, seconds = divmod(seconds, 60)
    return f"{int(hours):02}:{int(minutes):02}:{int(seconds):02}"

def write_to_file():
    global usage_records, stop_threads
    while not stop_threads:
        with usage_lock:
            if usage_records:
                try:
                    wb = openpyxl.load_workbook(log_file_path)
                    ws = wb["Session_Log"]

                    # Append current session data
                    for (app_name, app_title), duration in usage_records.items():
                        timestamp = datetime.now().strftime("%H:%M:%S %d-%m-%Y")
                        ws.append([app_name, app_title, format_duration(duration), timestamp])

                    # Auto-fit columns
                    for col in ws.columns:
                        max_length = max(len(str(cell.value)) for cell in col if cell.value)
                        col_letter = get_column_letter(col[0].column)
                        ws.column_dimensions[col_letter].width = max_length + 2

                    wb.save(log_file_path)
                    usage_records.clear()
                except Exception as e:
                    print(f"[Error] Writing to file: {e}")

            time.sleep(sleepingTime)




def aggregate_and_finalize():
    try:
        wb = openpyxl.load_workbook(log_file_path)
        consolidated_records = {}

        # Process the Session_Log sheet
        if "Session_Log" in wb.sheetnames:
            log_sheet = wb["Session_Log"]
            for row in log_sheet.iter_rows(min_row=2, values_only=True):
                app_name, app_title, duration, _ = row
                if app_name and app_title:
                    hours, minutes, seconds = map(int, duration.split(':'))
                    duration_in_seconds = hours * 3600 + minutes * 60 + seconds
                    key = (app_name, app_title)
                    consolidated_records[key] = consolidated_records.get(key, 0) + duration_in_seconds

        # Create or update Summary sheet
        if "Summary" not in wb.sheetnames:
            summary_sheet = wb.create_sheet(title="Summary")
        else:
            summary_sheet = wb["Summary"]
            summary_sheet.delete_rows(1, summary_sheet.max_row)

        summary_sheet.append(["App Name", "App Title", "Total Duration (HH:MM:SS)"])

        # Write consolidated data
        for (app_name, app_title), total_duration in consolidated_records.items():
            summary_sheet.append([app_name, app_title, format_duration(total_duration)])

        # Auto-fit columns
        for col in summary_sheet.columns:
            max_length = max(len(str(cell.value)) for cell in col if cell.value)
            col_letter = get_column_letter(col[0].column)
            summary_sheet.column_dimensions[col_letter].width = max_length + 2

        wb.save(log_file_path)
        print("[INFO] Aggregated data saved in 'Summary' sheet.")
    except Exception as e:
        print(f"[Error] Aggregating data: {e}")


def aggregate_detailed_usage():
    try:
        wb = openpyxl.load_workbook(log_file_path)

        # Consolidated records to track total time per app
        app_consolidated_records = {}

        # Process the Session_Log sheet
        if "Session_Log" in wb.sheetnames:
            log_sheet = wb["Session_Log"]

            # Iterate through log entries
            for row in log_sheet.iter_rows(min_row=2, values_only=True):
                app_name, app_title, duration, timestamp = row

                # Convert duration to seconds
                hours, minutes, seconds = map(int, duration.split(':'))
                duration_in_seconds = hours * 3600 + minutes * 60 + seconds

                # Create a unique key for each app and title combination
                key = (app_name, app_title)

                # If key doesn't exist, initialize the record
                if key not in app_consolidated_records:
                    app_consolidated_records[key] = {
                        'total_duration': duration_in_seconds,
                        'sessions': 1,
                        'first_seen': timestamp,
                        'last_seen': timestamp
                    }
                else:
                    # Update existing record
                    record = app_consolidated_records[key]
                    record['total_duration'] += duration_in_seconds
                    record['sessions'] += 1
                    record['last_seen'] = timestamp

        # Create or update Detailed_Summary sheet
        if "Detailed_Summary" not in wb.sheetnames:
            detailed_sheet = wb.create_sheet(title="Detailed_Summary")
        else:
            detailed_sheet = wb["Detailed_Summary"]
            detailed_sheet.delete_rows(1, detailed_sheet.max_row)

        # Write headers
        detailed_sheet.append([
            "App Name",
            "App Title",
            "Total Duration (HH:MM:SS)",
            "Number of Sessions",
            "Average Session Duration",
            "First Seen",
            "Last Seen"
        ])

        # Sort records by total duration in descending order
        sorted_records = sorted(
            app_consolidated_records.items(),
            key=lambda x: x[1]['total_duration'],
            reverse=True
        )

        # Write detailed records
        for (app_name, app_title), record in sorted_records:
            # Calculate average session duration
            avg_duration = record['total_duration'] / record['sessions']

            detailed_sheet.append([
                app_name,
                app_title,
                format_duration(record['total_duration']),
                record['sessions'],
                format_duration(avg_duration),
                record['first_seen'],
                record['last_seen']
            ])

        # Auto-fit columns
        for col in detailed_sheet.columns:
            max_length = max(len(str(cell.value)) for cell in col if cell.value)
            col_letter = get_column_letter(col[0].column)
            detailed_sheet.column_dimensions[col_letter].width = max_length + 2

        wb.save(log_file_path)
        print("[INFO] Detailed usage summary created.")

    except Exception as e:
        print(f"[Error] Creating detailed summary: {e}")



# Register the updated signal handler
signal.signal(signal.SIGINT, signal_handler)



def track_app_usage():
    last_app = None
    last_title = None
    start_time = None

    while not stop_threads:
        try:
            hwnd = win32gui.GetForegroundWindow()
            if hwnd:
                if win32gui.IsIconic(hwnd): # checks if the current window is minimized
                    print("current window is minimized. Putting the thread to sleep")
                    time.sleep(sleepingTime)
                    continue

                _, process_id = win32process.GetWindowThreadProcessId(hwnd)
                process = psutil.Process(process_id)
                current_app = process.name()
                current_title = truncate_title(win32gui.GetWindowText(hwnd))
            else:
                current_app, current_title = "Unknown", "No Active Window"

            current_time = time.time()

            # Initialize tracking for the first app
            if start_time is None:
                last_app = current_app
                last_title = current_title
                start_time = current_time
                continue

            # Detect app switch or long inactivity
            if current_app != last_app or current_title != last_title:
                # Log the app usage duration before switching
                duration = current_time - start_time
                if duration > 0:
                    print(f"[INFO] App: {last_app}, Title: {last_title}, Duration: {format_duration(duration)}")
                    log_usage(last_app, last_title, duration)  # Log raw seconds

                # Update the app details
                last_app = current_app
                last_title = current_title
                start_time = current_time

            time.sleep(sleepingTime)

        except Exception as e:
            print(f"[Error] {e}")
            time.sleep(sleepingTime)

    # Log final usage when thread stops
    if last_app and start_time:
        final_duration = time.time() - start_time
        if final_duration > 0:
            print(f"[INFO] Final App: {last_app}, Final Title: {last_title}, Final Duration: {format_duration(final_duration)}")
            log_usage(last_app, last_title, final_duration)


def log_usage(app_name, app_title, duration):
    print("detected app switch on process: "+ str(app_title) )
    global usage_records
    key = (app_name, app_title)
    with usage_lock:
        if key in usage_records:
            print("The Key" + str(key) + " is present in usage_records. Value is " + str(usage_records[key]))
            usage_records[key] += duration
            print("Value after addition: " + str(usage_records[key]))
        else:
            #print("The Key" + str(key) + " is not present in usage_records. Adding with duration " + str(duration))
            usage_records[key] = duration


def truncate_title(title, max_length=50):
    """
    Truncate the window title to a specified maximum length.
    If the title is longer, it will be cut and end with an ellipsis.

    :param title: Original window title
    :param max_length: Maximum allowed length of the title
    :return: Truncated title
    """
    if not title:
        return title

    if len(title) <= max_length:
        return title

    # Truncate and add ellipsis
    return title[:max_length-3] + "..."

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

# Ensure workbook is created at start
create_workbook()

# Register signal handler
signal.signal(signal.SIGINT, signal_handler)

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