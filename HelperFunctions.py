import json
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path

from dotenv import load_dotenv


# Get absolute path of a file or dir
def get_abs_path(path: str) -> str:
    project_root = os.path.abspath(os.path.join(os.path.dirname(__file__)))
    return os.path.join(project_root, path)


root_folder = get_abs_path('')
print(f'Root folder: {root_folder}')

today = datetime.now()
formatted_date = today.strftime('%d %b, %Y')

is_linux = sys.platform.startswith('linux')
if is_linux:
    today = today + timedelta(hours=5, minutes=30)
todayDate = today.strftime("%Y%m%d")
todayDateInDash = today.strftime("%Y-%m-%d")

today = datetime.now()

# Monday = 0
if today.weekday() == 0:   # Monday
    yesterday = today - timedelta(days=3)   # Go back to Friday
else:
    yesterday = today - timedelta(days=1)   # Normal yesterday

yesterdayDate = yesterday.strftime("%Y%m%d")
yesterdayDateInDash = yesterday.strftime("%Y-%m-%d")

# --- Current date + time ---
currentDatetime = datetime.now().strftime("%Y%m%d_%H%M%S")
currentDatetimeInDash = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# To detect weekends
detect_weekend = datetime.now().weekday()

# Set up logging
log_folder_path = os.path.join(root_folder, 'logs')
os.makedirs(log_folder_path, exist_ok=True)


def isProduction():
    return getattr(sys, 'frozen', False)


# set up Appsettings.json
appsettings_file_path = os.path.join(root_folder, 'Appsettings.json')
# logging.info(f"Appsettings file path: {appsettings_file_path}")
appsettings = json.loads(open(appsettings_file_path).read())


def get_appsettings(key):
    return appsettings.get(key)


FORCE_PRODUCTION = get_appsettings('FORCE_PRODUCTION')
IS_PRODUCTION = FORCE_PRODUCTION or isProduction()

