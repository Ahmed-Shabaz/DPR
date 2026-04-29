import time
from datetime import datetime
from win10toast import ToastNotifier
import sys
import traceback
import os

# --- Fix for icon path when converted to EXE ---
def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS  # PyInstaller temp path
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# Accepts input in either 'HH:MM' or 'HH:MM AM/PM' formats
def parse_time(timestr):
    try:
        return datetime.strptime(timestr, "%H:%M").time()  # 24-hour format
    except ValueError:
        return datetime.strptime(timestr, "%I:%M %p").time()  # 12-hour format


def is_notification_time(start_str, end_str):
    now = datetime.now()

    # Optional debug logs
    print(f"Current Time (12-hr): {now.strftime('%I:%M %p')}")
    print(f"Current Time (24-hr): {now.strftime('%H:%M')}")

    start_time = parse_time(start_str)
    end_time = parse_time(end_str)
    current_time = now.time()

    return start_time <= current_time <= end_time


def send_notification():
    try:
        print("Sending notification using win10toast...")

        icon_file = resource_path("DPR_2.ico")  # Load .ico from bundled EXE or dev path
        toaster = ToastNotifier()
        toaster.show_toast(
            " ",  # Blank title to hide "Python"
            "Submit the DPR by 7:30 PM. If it's already submitted, kindly ignore this message.",
            icon_path=icon_file,
            duration=5,
            threaded=True
        )
        print("Notification sent.\n")
    except Exception as e:
        print(f"Notification failed: {e}")
        traceback.print_exc()


def main():
    # Set your desired notification time window
    start_time_input = "10:34 AM"
    end_time_input = "07:30 PM"

    print("Notification checker started...\n")

    try:
        while True:
            if is_notification_time(start_time_input, end_time_input):
                send_notification()
                time.sleep(900)  # Notify every 15 minutes
            else:
                print("Not in the notification time window.\n")
                time.sleep(60)  # Check every 1 minute
    except KeyboardInterrupt:
        print("Notification checker stopped manually.")
    except Exception as e:
        print("An error occurred:", e)
        traceback.print_exc()
        time.sleep(5)
        sys.exit(1)


if __name__ == "__main__":
    main()
