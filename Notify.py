import time
from datetime import datetime
from plyer import notification
import sys
import traceback
import platform
import ctypes.wintypes  


# Accepts input in either 'HH:MM' or 'HH:MM AM/PM' formats
def parse_time(timestr):
    try:
        return datetime.strptime(timestr, "%H:%M").time()  # 24-hour format
    except ValueError:
        return datetime.strptime(timestr, "%I:%M %p").time()  # 12-hour format

def is_notification_time(start_str, end_str):
    now = datetime.now()

    # Optional debug logs (can remove later)
    print(f"Current Time (12-hr): {now.strftime('%I:%M %p')}")
    print(f"Current Time (24-hr): {now.strftime('%H:%M')}")

    start_time = parse_time(start_str)
    end_time = parse_time(end_str)
    current_time = now.time()

    return start_time <= current_time <= end_time

def send_notification():
    try:
        print("Sending notification using plyer...")
        notification.notify(
            title="Fill the DPR",
            message="Fill the DPR by 7:30 PM. If it is already filled, kindly ignore this message.",
            timeout=5
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
                time.sleep(900)  # Notify every 900 seconds in time window
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
