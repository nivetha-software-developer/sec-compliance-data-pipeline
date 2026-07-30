import re
import pandas as pd
import paramiko
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import sys
from HelperFunctions import root_folder, todayDate, currentDatetimeInDash, todayDateInDash, yesterdayDate, detect_weekend
from Utilities.Logger import Logger
import json
from datetime import datetime
from openpyxl import Workbook
import pdfkit
# from pypdf import PdfMerger
from pypdf import PdfWriter
from PyPDF2 import PdfMerger
import shutil
from dotenv import load_dotenv
import os
import SO_Validator
from postmarker.core import PostmarkClient

postmark = PostmarkClient(server_token='postmark_api_key')

config = pdfkit.configuration(wkhtmltopdf=r"path_wkhtmltopdf.exe")


env_file = os.path.join(root_folder, "env")
load_dotenv(dotenv_path=env_file)

SFTP_HOST = os.getenv("SFTP_HOST")
SFTP_USER = os.getenv("SFTP_USER")
SFTP_PASS = os.getenv("SFTP_PASS")
SFTP_PORT = int(os.getenv("SFTP_PORT"))
SFTP_BASE_PATH = os.getenv("SFTP_BASE_PATH")
# ========================
# FORM TYPES
# ========================
# DEF 14C, DEFC14A, 20-F, DEF 14A, F-3, S-3, 10-K, 10-Q
forms = ['DEF 14C', 'DEFC14A', '20-F', 'DEF 14A', 'F-3', 'S-3', '10-K', '10-Q']    # Totally " 8 " Form Types for SO
base_url = "https://www.sec.gov"

# ========================
# RUN COUNT & FOLDER SETUP
# ========================

base_output_dir = os.path.join(root_folder, "Data", todayDate)
os.makedirs(base_output_dir, exist_ok=True)

runcount_file = os.path.join(base_output_dir, "runcount.txt")

MAX_RUNS_PER_DAY = 10

# ----------------------------------
# STEP 1: Determine Current Run
# ----------------------------------

if os.path.exists(runcount_file):
    # Read existing count
    with open(runcount_file, "r") as f:
        try:
            last_run = int(f.read().strip())
        except ValueError:
            last_run = 0

    current_run = last_run + 1
else:
    # First run of the day
    current_run = 1

# ----------------------------------
# STEP 2: Stop Condition
# ----------------------------------

if current_run > MAX_RUNS_PER_DAY:
    logger(f"Maximum run limit reached for {todayDate}. ({MAX_RUNS_PER_DAY}/{MAX_RUNS_PER_DAY})", Logger.INFO)
    sys.exit(0)

# ----------------------------------
# Logger Condition
# ----------------------------------
logger_man = Logger(f'SEC_SO_{str(current_run)}', root_folder=root_folder, todayDate=todayDate)
logger, log_file_name = logger_man.log, logger_man.log_file_name

# ------------------------------------
# STEP 3: Create Entire Files & Folder
# ------------------------------------

run_folder = os.path.join(base_output_dir, str(current_run))

if not os.path.exists(run_folder):
    os.makedirs(run_folder)
else:
    logger(f"Warning: Folder {current_run} already exists.", Logger.WARNING)
    for filename in os.listdir(run_folder):
        file_path = os.path.join(run_folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)  # remove file
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)  # remove folder
        except Exception as e:
            logger(f"Failed to delete {file_path}. Reason: {e}", Logger.ERROR)

logger(f"Current Run Count: {current_run}", Logger.INFO)
logger(f"Output Folder: {run_folder}", Logger.INFO)

# ----------------------------
# Create Excel File
# ----------------------------

temp_pdf_download_dir = os.path.join(run_folder, "Temp")
os.makedirs(temp_pdf_download_dir, exist_ok=True)

excel_filename = f"SO{todayDate}_{current_run}.xlsx"
so_excel_path = os.path.join(run_folder, excel_filename)

wb = Workbook()
ws = wb.active
ws.title = "Sheet1"

# Headers
headers = ["Sno", "UniqueID", "Form Type", "URL", "Filename", "Remarks"]
ws.append(headers)
wb.save(so_excel_path)

temp_excel_file = os.path.join(run_folder, "sec_filings.xlsx")

scriptlog_path = os.path.join(run_folder, "scriptlog.txt")
scriptlog1_path = os.path.join(run_folder, "scriptlog1.txt")

url_path = os.path.join(run_folder, "url.txt")
url1_path = os.path.join(run_folder, "url1.txt")

note_file_path = os.path.join(run_folder, "Note.txt")
log_file_path = os.path.join(run_folder, "Log.txt")

input_json_path = os.path.join(run_folder, "input.json")
info_json_path = os.path.join(run_folder, "info.json")

actual_file_count = 0

input_structure = {
    "id": str(current_run),
    "datetime": currentDatetimeInDash,
    "type": "so",
    "Forms": {form: [] for form in forms}
}

with open(input_json_path, "w", encoding="utf-8") as input_json_file:
    json.dump(input_structure, input_json_file, indent=4)

with open(info_json_path, "w", encoding="utf-8") as info_json_file:
    json.dump(input_structure, info_json_file, indent=4)

logger("info & input json files created", Logger.INFO)

# Create Note.txt as empty
with open(note_file_path, "w", encoding="utf-8"):
    pass

# Create Log.txt as empty
with open(log_file_path, "w", encoding="utf-8"):
    pass
logger("note & log txt files created", Logger.INFO)

# ========================
# CHROME SETUP
# ========================

chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--disable-infobars")
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36"
)

try:
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
except Exception as e:
    logger(f"PermissionError: {e}", Logger.INFO)
    driver = webdriver.Chrome()

wait = WebDriverWait(driver, 50)

def get_previous_scriptlog_path():

    base_data_path = os.path.join(root_folder, "Data")

    # -------------------------------------
    # CASE 1: Same day, run > 1
    # -------------------------------------
    if current_run > 1:

        today_folder = os.path.join(base_data_path, todayDate)

        for run in range(current_run - 1, 0, -1):
            scriptlog_path = os.path.join(
                today_folder,
                str(run),
                "scriptlog1.txt"
            )

            if os.path.exists(scriptlog_path):
                return scriptlog_path

    # -------------------------------------
    # CASE 2: Run 1 → check previous date
    # -------------------------------------
    else:

        all_dates = [
            d for d in os.listdir(base_data_path)
            if d.isdigit()
        ]

        all_dates = sorted(all_dates)

        if todayDate not in all_dates:
            return None

        today_index = all_dates.index(todayDate)

        if today_index == 0:
            return None  # no previous date exists

        previous_date = all_dates[today_index - 1]
        previous_date_folder = os.path.join(base_data_path, previous_date)

        runcount_file = os.path.join(previous_date_folder, "runcount.txt")

        if not os.path.exists(runcount_file):
            return None

        with open(runcount_file, "r") as f:
            last_run = f.read().strip()

        scriptlog_path = os.path.join(
            previous_date_folder,
            last_run,
            "scriptlog1.txt"
        )

        if os.path.exists(scriptlog_path):
            return scriptlog_path

    return None

def to_datetime_safe(date_str, time_str):
    try:
        return datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M:%S")
    except:
        return None
        
def process_incremental_from_scriptlog1():
    base_data_path = os.path.join(root_folder, "Data")
    today_folder = os.path.join(base_data_path, todayDate)

    # ---------------------------------------------------
    # STEP 1: Get previous scriptlog1.txt
    # ---------------------------------------------------

    prev_scriptlog = get_previous_scriptlog_path()

    if not prev_scriptlog or not os.path.exists(prev_scriptlog):
        logger("No previous scriptlog1.txt found.", Logger.INFO)
        return

    with open(prev_scriptlog, "r", encoding="utf-8") as f:
        log_lines = [line.strip() for line in f if line.strip()]

    if not log_lines:
        logger("Previous scriptlog1.txt empty.", Logger.INFO)
        return

    # ---------------------------------------------------
    # STEP 2: Load Excel
    # ---------------------------------------------------

    if not os.path.exists(temp_excel_file):
        logger("Excel not found.", Logger.INFO)
        return

    # Remove specified form types before reading
    temp_df = pd.read_excel(temp_excel_file)
    forms_to_remove = ['F-3ASR', 'S-3ASR', 'S-3DPOS']
    temp_df = temp_df[~temp_df['Form'].isin(forms_to_remove)]
    temp_df.to_excel(temp_excel_file, index=False)
    
    df = pd.read_excel(temp_excel_file)

    df["Form"] = df["Form"].astype(str).str.strip()
    df["Filing_Date"] = df["Filing_Date"].astype(str).str.strip()
    df["Accepted"] = df["Accepted"].astype(str).str.strip()
    df["Accepted"] = df["Accepted"].str.split(".").str[0]

    all_above_rows = pd.DataFrame()

    # Store unmatched previous timestamps
    unmatched_previous = {}

    # ---------------------------------------------------
    # STEP 3: Check each previous log entry
    # ---------------------------------------------------
    checked_entries = []

    for line in log_lines:

        if " - " not in line:
            continue

        left_part, form_type = line.split(" - ")
        form_type = form_type.strip()

        cleaned = re.findall(r"\d{4}-\d{2}-\d{2}|\d{2}:\d{2}:\d{2}", left_part)

        if len(cleaned) < 3:
            continue

        filing_date = cleaned[0]
        accepted_time = cleaned[2]

        logger(f"Checking Form: {form_type}", Logger.SUCCESS)
        logger(f"Checking Timestamp: {filing_date} {accepted_time}", Logger.SUCCESS)

        # 🔥 Save for fallback logging
        checked_entries.append((form_type, filing_date, accepted_time))

        form_df = df[df["Form"].str.startswith(form_type, na=False)]

        if form_df.empty:
            unmatched_previous[form_type] = f"{filing_date} {filing_date} {accepted_time}"
            continue

        form_df = form_df.reset_index(drop=True)

        match_index = None

        for i, row in form_df.iterrows():

            excel_date = re.findall(r"\d{4}-\d{2}-\d{2}", str(row["Filing_Date"]))
            excel_time = re.findall(r"\d{2}:\d{2}:\d{2}", str(row["Accepted"]))

            if not excel_date or not excel_time:
                continue

            if excel_date[0] == filing_date and excel_time[0] == accepted_time:
                match_index = i
                break

        # ---------------------------------------------------
        # NEW FALLBACK LOGIC : If exact timestamp not found
        # ---------------------------------------------------
        if match_index is None:

            logger(f"No exact match found. Checking for newer filings for {form_type}", Logger.INFO)

            # Convert scriptlog timestamp to datetime
            prev_dt = to_datetime_safe(filing_date, accepted_time)

            if prev_dt is None:
                unmatched_previous[form_type] = f"{filing_date} {filing_date} {accepted_time}"
                continue

            # Collect rows GREATER than scriptlog timestamp
            newer_rows = []

            for i, row in form_df.iterrows():

                excel_date = re.findall(r"\d{4}-\d{2}-\d{2}", str(row["Filing_Date"]))
                excel_time = re.findall(r"\d{2}:\d{2}:\d{2}", str(row["Accepted"]))

                if not excel_date or not excel_time:
                    continue

                excel_dt = to_datetime_safe(excel_date[0], excel_time[0])

                if excel_dt and excel_dt > prev_dt:
                    newer_rows.append(row)

            # If newer filings exist → treat as incremental data
            if newer_rows:
                logger(f"New filings detected for {form_type} (timestamp mismatch case)", Logger.SUCCESS)

                newer_df = pd.DataFrame(newer_rows)
                all_above_rows = pd.concat([all_above_rows, newer_df])

                # Update scriptlog with latest datetime from Excel
                latest_row = newer_df.iloc[0]
                latest_date = re.findall(r"\d{4}-\d{2}-\d{2}", str(latest_row["Filing_Date"]))[0]
                latest_time = re.findall(r"\d{2}:\d{2}:\d{2}", str(latest_row["Accepted"]))[0]

                ordered_results = {} if "ordered_results" not in locals() else ordered_results
                ordered_results[form_type] = f"{latest_date} {latest_date} {latest_time}"

                continue

            # If still nothing newer → keep previous timestamp
            unmatched_previous[form_type] = f"{filing_date} {filing_date} {accepted_time}"
            continue

        above_rows = form_df.iloc[:match_index]

        if not above_rows.empty:
            all_above_rows = pd.concat([all_above_rows, above_rows])

    # ---------------------------------------------------
    # STEP 4: Remove duplicates
    # ---------------------------------------------------

    if all_above_rows.empty:

        logger("No new rows found above matches.", Logger.INFO)

        with open(scriptlog_path, "a", encoding="utf-8") as slog, \
                open(scriptlog1_path, "a", encoding="utf-8") as slog1:

            for form_type, filing_date, accepted_time in checked_entries:
                timestamp = f"{filing_date} {filing_date} {accepted_time}"

                slog.write(timestamp + "\n")
                slog1.write(f"{timestamp} - {form_type}\n")

        logger("Checked timestamps appended to scriptlog files.", Logger.SUCCESS)

    # ---------------------------------------------------
    # STEP 5: Write url.txt & url1.txt
    # ---------------------------------------------------

    all_above_rows = all_above_rows.drop_duplicates().reset_index(drop=True)

    with open(url_path, "w", encoding="utf-8") as f1, \
            open(url1_path, "w", encoding="utf-8") as f2:

        for _, row in all_above_rows.iterrows():

            html_link = str(row["HTML_Link"]).strip()
            filing_val = str(row["Filing_Date"]).strip()
            company_name = str(row["Company_Name"]).strip()
            accession_num = str(row["Accession_Number"]).strip()
            act = str(row["Act"]).strip()
            file_size = str(row["Size"]).strip()

            accepted_val = (
                str(row["Accepted"])
                .replace("\n", " ")
                .replace("\r", " ")
                .strip()
            )

            if len(accepted_val) == 8:
                accepted_datetime = f"{filing_val} {accepted_val}"
            else:
                accepted_datetime = accepted_val

            f1.write(html_link + "\n")

            f2.write(f"{row['Form']} | {company_name} | {html_link} | {accession_num} | {act} | {file_size} | {accepted_datetime} | {filing_val}\n")

    # ---------------------------------------------------
    # STEP 6: Collect FIRST row per allowed form
    # ---------------------------------------------------

    ordered_results = {}

    # First add unmatched previous timestamps
    for form in forms:
        if form in unmatched_previous:
            ordered_results[form] = unmatched_previous[form]

    try:
        # Then add new Excel rows (first occurrence per form)
        for form in forms:

            form_rows = all_above_rows[
                all_above_rows["Form"].str.startswith(form, na=False)
            ]

            if form_rows.empty:
                continue

            first_row = form_rows.iloc[0]

            filing_val = str(first_row["Filing_Date"]).strip()
            accepted_val = " ".join(str(first_row["Accepted"]).split())

            if len(accepted_val) == 8:
                accepted_datetime = f"{filing_val} {accepted_val}"
            else:
                accepted_datetime = accepted_val

            ordered_results[form] = f"{filing_val} {accepted_datetime}"

        # ---------------------------------------------------
        # FIX: Ensure all forms retain previous timestamp
        # ---------------------------------------------------

        previous_map = {
            form_type: f"{filing_date} {filing_date} {accepted_time}"
            for form_type, filing_date, accepted_time in checked_entries
        }

        for form in forms:
            if form not in ordered_results and form in previous_map:
                ordered_results[form] = previous_map[form]

        # ---------------------------------------------------
        # STEP 7: Write scriptlogs in STRICT forms order
        # ---------------------------------------------------

        with open(scriptlog_path, "w", encoding="utf-8") as slog, \
                open(scriptlog1_path, "w", encoding="utf-8") as slog1:

            for form in forms:
                if form in ordered_results:
                    timestamp = ordered_results[form]

                    slog.write(timestamp + "\n")
                    slog1.write(f"{timestamp} - {form}\n")

        logger("Complete flow executed in the form order.", Logger.SUCCESS)
    except KeyError:
        logger("KeyError Occurred for FORM.", Logger.INFO)
    except Exception as e:
        logger(f"Exception occurred in process_incremental_from_scriptlog1 function: {e}")

def normalize_form_type(form_type):
    if not form_type:
        return ""

    original = form_type.upper().strip()

    # Detect amendment before cleaning
    is_amendment = False

    # If contains /A (like 10-K/A)
    if "/A" in original:
        is_amendment = True

    # If ends with A (like DEFC14A, S-3A)
    if re.search(r'A$', original):
        is_amendment = True

    # If contains ASR (like S-3ASR)
    if "ASR" in original:
        is_amendment = True

    # Remove special characters (keep letters & numbers)
    cleaned = re.sub(r'[^A-Z0-9]', '', original)

    # Remove trailing A if amendment
    if is_amendment and cleaned.endswith("A"):
        cleaned = cleaned[:-1]

    # Remove ASR completely
    cleaned = cleaned.replace("ASR", "")

    # Final Amendment naming
    if is_amendment:
        cleaned = cleaned + "Amendment"

    return cleaned


def generate_pdf_filename(company_name, form_type, output_dir):

    # Remove (CIK) and anything inside brackets
    company_name = re.sub(r"\s*\(.*?\)", "", company_name)

    # # Remove special characters except letters & numbers
    company_name = re.sub(r'[^A-Za-z0-9&\-\.\s]', '', company_name)

    # # Replace spaces with underscore
    # company_name = company_name.strip().replace(" ", "_")

    # Convert to Sentence Case
    company_name = company_name.title()

    # Ensure no trailing dot/comma/underscore
    company_name = company_name.rstrip("._, ")

    # Clean form_type (already normalized ideally)
    form_type = re.sub(r'[^A-Za-z0-9]', '', form_type)

    base_name = f"{company_name}_{form_type}"

    # Start numbering from 1
    count = 1
    while True:
        filename = f"{base_name}_{count}.pdf"
        full_path = os.path.join(output_dir, filename)

        if not os.path.exists(full_path):
            return filename, full_path

        count += 1

def resolve_sftp_target_runcount(local_run, run_folder_name):
    """
    Mapping:

    1   → 10 (Yesterday)
    1A  → 11 (Yesterday)
    2   → 1  (Today)
    3   → 2  (Today)
    4   → 3  (Today)
    ...
    """

    # CASE 1A
    if run_folder_name == "1A":
        return "11", yesterdayDate
        
    # CASE 1 → 1
    if str(local_run) == "1":
        return "10", yesterdayDate

    # CASE 2+
    # Shift down by 1
    sftp_number = int(local_run) - 1
    return str(sftp_number), todayDate

def should_skip_sftp(detect_weekend, sftp_folder):
    # Sunday → block everything
    if detect_weekend == 6:
        return True

    # Saturday → only allow 10 and 11
    if detect_weekend == 5:
        return sftp_folder not in ["10", "11"]

    # Monday → block 10 and 11
    if detect_weekend == 0:
        return sftp_folder in ["10", "11"]

    # Tue-Fri → allow everything
    return False
    
def upload_to_sftp(run_folder):

    run_folder_name = os.path.basename(run_folder)

    sftp_folder, target_date = resolve_sftp_target_runcount(
        current_run,
        run_folder_name
    )

    logger(f"SFTP Folder: {sftp_folder}", Logger.INFO)
    logger(f"SFTP Date Folder: {target_date}", Logger.INFO)

    try:
        transport = paramiko.Transport((SFTP_HOST, SFTP_PORT))
        transport.connect(username=SFTP_USER, password=SFTP_PASS)

        sftp = paramiko.SFTPClient.from_transport(transport)

        if should_skip_sftp(detect_weekend, sftp_folder):
            logger_man.log(
                "Weekend detected. Skipping SFTP folder creation and upload.",
                Logger.INFO
            )
            sftp.close()
            transport.close()
            return
    
        sftp.chdir(SFTP_BASE_PATH)

        try:
            sftp.chdir(target_date)
        except IOError:
            sftp.mkdir(target_date)
            sftp.chdir(target_date)

        # Create run folder
        try:
            sftp.mkdir(sftp_folder)
        except:
            pass

        sftp.chdir(sftp_folder)

        # Upload Required Files
        for file_name in os.listdir(run_folder):

            if (
                file_name.endswith(".pdf")
                or file_name.endswith(".json")
                or file_name in ["Log.txt", "Note.txt", "url.txt", "scriptlog.txt"]
                or (file_name.startswith("SO") and file_name.endswith(".xlsx"))
            ):
                local_path = os.path.join(run_folder, file_name)
                # Default remote name
                remote_file_name = file_name

                # If this is the SO Excel file → rename it
                if file_name.startswith("SO") and file_name.endswith(".xlsx"):
                    remote_file_name = f"SO{target_date}_{sftp_folder}.xlsx"

                if detect_weekend == 6:  # Sunday
                    logger_man.log("Sunday detected. SFTP upload stopped.", Logger.INFO)
                    # sys.exit()

                elif detect_weekend == 5:  # Saturday
                    if sftp_folder in ["10", "11"]:
                        logger_man.log("Saturday allowed run (10/11). Starting SFTP upload...", Logger.INFO)
                        sftp.put(local_path, remote_file_name)
                        logger_man.log("File uploaded successfully.", Logger.SUCCESS)
                    else:
                        logger_man.log(f"Saturday run {sftp_folder} blocked. Only 10 and 11 allowed.", Logger.INFO)
                        # sys.exit()
                        
                elif detect_weekend == 0:  # Monday
                    if sftp_folder in ["10", "11"]:
                        logger_man.log("Monday not allowed run (10/11). Starting SFTP upload...", Logger.INFO)
                    else:
                        sftp.put(local_path, remote_file_name)
                        logger_man.log("Monday File uploaded successfully.", Logger.SUCCESS)

                else:
                    # Tuesday – Friday
                    logger_man.log("Weekday detected. Starting SFTP upload...", Logger.SUCCESS)
                    sftp.put(local_path, remote_file_name)
                    logger_man.log("File uploaded successfully [In ELSE].", Logger.SUCCESS)
                    
        sftp.close()
        transport.close()

    except Exception as e:
        logger(f"SFTP Upload Failed: {e}", Logger.ERROR)
        raise

def ensure_previous_day_folders(sftp, previous_sftp_date_path):

    try:
        sftp.chdir(previous_sftp_date_path)
    except Exception:
        logger("Previous SFTP date folder not found", Logger.INFO)
        return

    # Loop 1 to 9
    for folder_num in range(1, 10):
    
        folder_path = f"{SFTP_BASE_PATH}{yesterdayDate}/{str(folder_num)}"

        try:
            # Check if folder exists
            sftp.chdir(folder_path)
            logger(f"Folder exists in SFTP: {folder_num}", Logger.INFO)

        except Exception:

            logger(f"Missing folder in SFTP: {folder_num}", Logger.ERROR)

            # Create missing folder
            try:
                sftp.mkdir(folder_path)
                logger(f"Created folder in SFTP: {folder_num}", Logger.SUCCESS)
            except Exception as e:
                logger(f"Failed to create folder {folder_num}: {e}", Logger.ERROR)
                continue

            # ==========================================
            # Read scriptlog from previous folder (n-1)
            # ==========================================
            previous_folder_num = folder_num - 1

            if previous_folder_num <= 0:
                logger(f"No previous folder available for folder {folder_num}", Logger.WARNING)
                continue

            source_log_path = f"{SFTP_BASE_PATH}{yesterdayDate}/{str(previous_folder_num)}/scriptlog.txt"  
            target_log_path = f"{folder_path}/scriptlog.txt"

            try:
                # Read previous folder scriptlog
                with sftp.open(source_log_path, "r") as src:
                    content = src.read()

                # Write to new folder
                with sftp.open(target_log_path, "w") as dst:
                    dst.write(content)

                logger(
                    f"Copied scriptlog from folder {previous_folder_num} to {folder_num}",
                    Logger.SUCCESS
                )

                # Run validator for yesterday missing folder
                SO_Validator.validator_check_for_yesterday_path(
                    yesterdayDate,
                    folder_num,
                    folder_path,
                    currentDatetimeInDash,
                    forms,
                    sftp
                )

            except Exception as e:
                logger(
                    f"Failed to copy scriptlog from folder {previous_folder_num} "
                    f"to {folder_num}: {e}",
                    Logger.ERROR
                )

def check_missing_file_validation(run_folder_path):

    logger(f"Validating folder: {run_folder_path}", Logger.INFO)

    # -----------------------------
    # 1. Find SO*.xlsx file
    # -----------------------------
    so_excel = None

    for file in os.listdir(run_folder_path):
        if file.startswith("SO") and file.endswith(".xlsx"):
            so_excel = os.path.join(run_folder_path, file)
            break

    if not so_excel:
        logger("SO Excel file not found", Logger.ERROR)
        return False

    # -----------------------------
    # 2. Read Excel
    # -----------------------------
    try:
        df = pd.read_excel(so_excel)
        if "Filename" not in df.columns:
            print("Column 'Filename' not found in Excel")
            return False

        excel_filenames = (
            df["Filename"]
            .dropna()
            .astype(str)
            .str.strip()
            .tolist()
        )
    except Exception as e:
        logger(f"Failed to read Excel: {e}", Logger.ERROR)
        return False

    # Convert all values to string for safe comparison
    excel_values = df.astype(str).values.flatten()
    excel_values = [str(v).strip() for v in excel_values if str(v) != "nan"]

    # -----------------------------
    # 3. Collect downloaded PDFs
    # -----------------------------
    pdf_files = [
        f for f in os.listdir(run_folder_path)
        if f.lower().endswith(".pdf")
    ]

    # -----------------------------
    # 4. Compare
    # -----------------------------
    missing_files = []

    for excel_name in excel_filenames:

        if excel_name not in pdf_files:
            missing_files.append(excel_name)
        
    for pdf in pdf_files:
        pdf_name_without_ext = os.path.splitext(pdf)[0]

        # Check if PDF name appears anywhere in Excel
        found = any(pdf_name_without_ext in cell for cell in excel_values)

        if not found:
            missing_files.append(pdf)
    
    # -----------------------------
    # 5. Result
    # -----------------------------
    if missing_files:
        logger("Missing PDF entries in Excel:", Logger.INFO)
        for m in missing_files:
            logger(f"   -{m}", Logger.INFO)
                    
            # Get URL from Excel
            row = df[df["Filename"] == m]

            if not row.empty:
                url = row.iloc[0]["URL"]

                logger(f"     URL: {url}", Logger.INFO)
                SO_Validator.download_missing_pdfs(driver, url, run_folder_path, m, config, temp_pdf_download_dir)
        return True
    else:
        logger("All PDFs validated successfully", Logger.SUCCESS)
        return True

def check_thousands_in_txt(txt_url):

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/145.0.0.0 Safari/537.36 Edg/145.0.0.0",
        "Accept-Encoding": "gzip, deflate",
        "Host": "www.sec.gov"
    }
    try:
        response = requests.get(txt_url,headers, timeout=30)
        if response.status_code == 200:
            content = response.text.lower()
            if "thousand" in content or "thousands" in content:
                return "TRUE"
        return ""
    except Exception as e:
        logger(f"Error reading TXT file: {e}", Logger.ERROR)
        return ""

def can_send_mail():
    now = datetime.now()
    day = now.weekday()   # Monday=0, Sunday=6
    hour = now.hour

    # Saturday afternoon onwards
    if day == 5 and hour >= 12:
        return False

    # Sunday full day
    if day == 6:
        return False

    # Monday morning
    if day == 0 and hour < 12:
        return False

    return True
    
    
def send_mail():       
    if not can_send_mail():
        logger("Email skipped due to restricted time", Logger.INFO)
        return

    try:
        postmark.emails.send(
            From='XYZ@gmail.com',
            To='XXX@gmail.com',
            Subject='SEC_SO - LC [server_name] || Success!',
            HtmlBody=f"""
            The SFTP operation was completed successfully.<br><br>
            The folder has been successfully moved and no issues were encountered.<br><br>
            <b>SO :</b> {current_run}
            """,
            Attachments=[log_file_name]
        )

        logger("Mail sent successfully", Logger.SUCCESS)

    except Exception as e:
        logger(f"Mail sending failed: {e}", Logger.ERROR)
    

def generate_pdf_with_retry(url, output_path, config, retries=3, delay=10):
    for attempt in range(1, retries + 1):
        try:
            pdfkit.from_url(url, output_path, configuration=config)
            return True  # success

        except Exception as e:
            logger(f"[Attempt {attempt}] PDF generation failed for {url} → {str(e)}", Logger.WARNING)

            if attempt < retries:
                time.sleep(delay)
            else:
                logger(f"All retries failed for {url}", Logger.ERROR)
                return False
                
# ============================================================
# PRE-RUN SFTP CHECK (ONLY WHEN RUN = 1)
# ============================================================

if str(current_run) == "1":
    try:
        logger("Run count is 1 → Checking previous day SFTP folders...", Logger.INFO)

        transport = paramiko.Transport((SFTP_HOST, SFTP_PORT))
        transport.connect(username=SFTP_USER, password=SFTP_PASS)
        sftp = paramiko.SFTPClient.from_transport(transport)

        previous_sftp_date_path = os.path.join(SFTP_BASE_PATH, yesterdayDate)

        ensure_previous_day_folders(
            sftp,
            previous_sftp_date_path
        )

        sftp.close()
        transport.close()

        logger("Previous day SFTP folder check completed.", Logger.SUCCESS)

    except Exception as e:
        logger(f"SFTP pre-check failed: {e}", Logger.ERROR)
        raise  # Stop execution

# ============================================================
# ======================== MAIN TRY ==========================
# ============================================================

try:
    data = []

    for form in forms:
        encoded_form = form.replace(" ", "+")  # important for forms like DEF 14A
        start = 0
        count = 100

        stop_form_processing = False
        latest_filing_date = None
        latest_accepted_dt = None

        while True:

            paginated_url = (
                f"https://www.sec.gov/cgi-bin/browse-edgar?"
                f"action=getcurrent&company=&type={encoded_form}"
                f"&owner=include&start={start}&count={count}"
            )

            driver.get(paginated_url)

            try:
                table = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, "/html/body/div/table[2]"))
                )

                rows = table.find_elements(By.XPATH, ".//tr")

                if len(rows) <= 1:
                    logger("No rows found. Moving to next form.", Logger.INFO)
                    break

                company_name = None  # store company name for next filing row

                # ==========================
                # EXTRACT TABLE DATA
                # ==========================
                for row in rows:
                    cols = row.find_elements(By.TAG_NAME, "td")

                    # ----------------------------------------
                    # CASE 1: Company Name Row
                    # ----------------------------------------
                    # Company row has empty first two columns
                    if len(cols) >= 3 and cols[0].text.strip() == "" and cols[1].text.strip() == "":
                        clear_company_name = cols[2].text.strip()
                        company_name = re.sub(r"\s*\(.*?\)", "", clear_company_name).strip()
                        continue  # move to next row

                    if len(cols) < 6:
                        continue

                    form_text = cols[0].text.strip()
                    if not form_text:
                        continue

                    accepted_text = cols[3].text.strip().replace("\n", " ")
                    filing_date = cols[4].text.strip()

                    links = cols[1].find_elements(By.TAG_NAME, "a")
                    html_link = links[0].get_attribute("href") if len(links) > 0 else None
                    txt_link = links[1].get_attribute("href") if len(links) > 1 else None
                    description_raw = cols[2].text.strip().replace("\n", " ")
                    file_number = cols[5].text.strip().replace("\n", " ")

                    # Extract Accession Number
                    accession_match = re.search(r"Accession Number:\s*([^\s]+)", description_raw)
                    accession_number = accession_match.group(1) if accession_match else None

                    # Extract Act
                    act_match = re.search(r"Act:\s*([^\s]+)", description_raw)
                    act_value = act_match.group(1) if act_match else None

                    # Extract Size
                    size_match = re.search(r"Size:\s*([^\s]+\s*(KB|MB|GB))", description_raw)
                    size_value = size_match.group(1) if size_match else None

                    # Extract Main Description (before Accession Number)
                    main_description = description_raw.split("Accession Number:")[0].strip()

                    record = {
                        "Form": form_text,
                        "HTML_Link": html_link,
                        "TXT_Link": txt_link,
                        "Company_Name": company_name,
                        "Main_Description": main_description,
                        "Accession_Number": accession_number,
                        "Act": act_value,
                        "Size": size_value,
                        "Accepted": accepted_text,
                        "Filing_Date": filing_date,
                        "File_No": file_number
                    }

                    data.append(record)

                logger(f"Page extracted for form: {form} (start={start})", Logger.SUCCESS)

                # ==========================
                # CHECK NEXT 100 BUTTON
                # ==========================

                next_buttons = driver.find_elements(By.XPATH, "//input[@value='Next100']")
                if next_buttons:
                    start += count
                    time.sleep(1)
                else:
                    logger("No 'Next 100' button found. Pagination complete.", Logger.INFO)
                    break


            except Exception as e:
                logger(f"Error in form {form}: {str(e)}", Logger.ERROR)
                break

    # ========================
    # SAVE EXCEL
    # ========================

    if data:
        df = pd.DataFrame(data)
        before_count = len(df)
        df.drop_duplicates(inplace=True)
        after_count = len(df)
        df.reset_index(drop=True, inplace=True)
        df.to_excel(temp_excel_file, index=False)
        logger(
            f"Excel saved at: {temp_excel_file} | "
            f"Removed {before_count - after_count} duplicate rows",
            Logger.SUCCESS
        )
    else:
        logger("No data to save in Excel.", Logger.WARNING)

    process_incremental_from_scriptlog1()

    records = []

    with open(url1_path, "r", encoding="utf-8") as file:
        for line in file:
            if line.strip():
                parts = [p.strip() for p in line.split("|")]

                if len(parts) >= 8:
                    record = {
                        "form_type": parts[0],
                        "company_name": parts[1],
                        "url": parts[2],
                        "accession_no": parts[3],
                        "document_count": parts[4],
                        "file_size": parts[5],
                        "accepted_datetime": parts[6],
                        "filing_date": parts[7]
                    }

                    records.append(record)

    # ----------------------------
    # Load existing JSON
    # ----------------------------
    if os.path.exists(input_json_path):
        with open(input_json_path, "r", encoding="utf-8") as f:
            input_json_data = json.load(f)
    else:
        input_json_data = {}

    if os.path.exists(info_json_path):
        with open(info_json_path, "r", encoding="utf-8") as info_f:
            info_json_data = json.load(info_f)
    else:
        info_json_data = {}

    # ----------------------------
    # Append Records
    # ----------------------------
    for record in records:

        raw_type = record["form_type"].strip().upper()

        # Clean only for comparison (remove space, hyphen, slash)
        compare_type = raw_type.replace("-", "").replace(" ", "").replace("/", "")

        matched_key = None

        for key in input_json_data["Forms"].keys():

            key_compare = key.upper().replace("-", "").replace(" ", "").replace("/", "")

            # 1. Exact match first
            if compare_type == key_compare:
                matched_key = key
                break

            # 2. Startswith match (S-3ASR, S-3DPOS, etc.)
            if compare_type.startswith(key_compare):
                matched_key = key
                break

        # Skip if no matching form family found
        if not matched_key:
            continue

        # Split accepted datetime
        accepted_date = ""
        accepted_time = ""

        if record["accepted_datetime"]:
            parts = record["accepted_datetime"].split()
            if len(parts) >= 2:
                accepted_date = parts[0]
                accepted_time = parts[1]

        # Prepare new entry
        new_entry = {
            "url": record["url"],
            "name": record["company_name"],
            "filesize": record["file_size"],
            "filing date": record["filing_date"],
            "accepted date": accepted_date,
            "accepted time": accepted_time
        }


        # Append into correct form array
        input_json_data["Forms"][matched_key].append(new_entry)

    # ----------------------------
    # Save Back to JSON
    # ----------------------------
    with open(input_json_path, "w", encoding="utf-8") as f:
        json.dump(input_json_data, f, indent=4)

    logger("Records successfully appended to input.json", Logger.SUCCESS)

    # ========================
    # LOOP THROUGH RECORDS
    # ========================

    with open(log_file_path, "a", encoding="utf-8") as log_file:
        for index, record in enumerate(records, start=1):
            logger(f"Opening Record {index}", Logger.INFO)

            normalized_form_type = normalize_form_type(record["form_type"])

            print("Form Type:", record["form_type"])
            print("Company:", record["company_name"])

            driver.get(record["url"])
            doc_table = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="contentDiv"]/div[2]')))

            rows = doc_table.find_elements(By.TAG_NAME, "tr")[1:]
            document_links = []
            for row in rows:
                try:

                    cols = row.find_elements(By.TAG_NAME, "td")

                    # Skip header or invalid rows
                    if len(cols) < 4:
                        continue

                    # Get Type column (4th column)
                    type_text = cols[3].text.strip()

                    link_element = row.find_element(By.TAG_NAME, "a")
                    href = link_element.get_attribute("href")

                    link_element = row.find_element(By.TAG_NAME, "a")
                    link_text = link_element.text.lower()

                    if link_text.endswith((".htm", ".html")):
                        href_val = link_element.get_attribute("href")

                        if href_val and not href_val.startswith(("http", "https")):
                            href_val = base_url + href_val

                        if href_val:
                            document_links.append(href_val)

                    # ------------------------------------------
                    # If .txt → check for thousands word
                    # ------------------------------------------
                    if href.lower().endswith(".txt"):
                        if type_text:  # ✅ Not empty (NOT &nbsp;)
                            if href and not href.startswith(("http", "https")):
                                href_val = base_url + href
                            if href_val:
                                document_links.append(href_val)

                        else:
                            result = check_thousands_in_txt(href)
                            if result == "TRUE":
                                remarks_value = "thousand/thousands are available"
                            else:
                                remarks_value = ""

                except Exception as e:
                    logger(f"Error while indentify the document: {e}", Logger.ERROR)
                    continue
            logger(f"List of Document Links are:\n {document_links}", Logger.SUCCESS)

            # ----------------------------
            # Unique ID generating
            # ----------------------------
            now = datetime.now()
            unique_id_val = ("SO" + now.strftime("%y%m%d%H%M%S") + now.strftime("%f")[:8])
            pdf_name, pdf_path = generate_pdf_filename(
                record["company_name"],
                normalized_form_type,
                run_folder
            )

            ws.append([
                index,                  # Sno
                unique_id_val,          # UniqueID
                record["form_type"],    # Form Type
                record["url"],          # URL
                pdf_name,               # Filename
                remarks_value           # Remarks
            ])
            wb.save(so_excel_path)
            # ===============================
            # CASE 1: Only One HTML
            # ===============================
            if len(document_links) == 1:
                time.sleep(.2)
                logger("Single document found. Converting to PDF...", Logger.INFO)
                
                clean_url = document_links[0].replace("ix?doc=/", "")
                url_to_download = clean_url if clean_url else document_links[0]
                success = generate_pdf_with_retry(url_to_download, pdf_path, config)

                if not success:
                    logger(f"Failed to generate PDF: {url_to_download}", Logger.ERROR)

            # ===============================
            # CASE 2: Multiple HTML → Merge
            # ===============================
            elif len(document_links) > 1:
                time.sleep(.2)
                logger("Multiple documents found. Merging into single PDF...", Logger.INFO)

                merger = PdfMerger()
                # merger = PdfWriter()
                temp_pdfs = []

                for i, url in enumerate(document_links):
                    temp_pdf = os.path.join(temp_pdf_download_dir, f"temp_{i}.pdf")
                    clean_url = url.replace("ix?doc=/", "")
                    url_to_download = clean_url if clean_url else url
                    success = generate_pdf_with_retry(url_to_download, temp_pdf, config)

                    if success:
                        merger.append(temp_pdf)
                        temp_pdfs.append(temp_pdf)
                    else:
                        logger(f"Skipping failed document: {url_to_download}", Logger.WARNING)

                if temp_pdfs:
                    merger.write(pdf_path)
                    merger.close()
                else:
                    logger("No PDFs were successfully created to merge.", Logger.ERROR)
                # Remove temp PDFs
                for temp in temp_pdfs:
                    os.remove(temp)

            logger(f"Final PDF created at: {pdf_path}", Logger.SUCCESS)

            new_info_entry = {
                "id": unique_id_val,
                "url": record["url"],
                "pdf": pdf_name
            }

            raw_type = record["form_type"].strip().upper()

            # Clean only for comparison (remove space, hyphen, slash)
            compare_type = raw_type.replace("-", "").replace(" ", "").replace("/", "")

            matched_key = None

            for key in input_json_data["Forms"].keys():

                key_compare = key.upper().replace("-", "").replace(" ", "").replace("/", "")

                # 1️⃣ Exact match first
                if compare_type == key_compare:
                    matched_key = key
                    break

                # 2️⃣ Startswith match (S-3ASR, S-3DPOS, etc.)
                if compare_type.startswith(key_compare):
                    matched_key = key
                    break

            # Skip if no matching form family found
            if not matched_key:
                continue

            # Append into correct form array
            info_json_data["Forms"][matched_key].append(new_info_entry)

            # ============================
            # LOG FILE STRUCTURE
            # ============================

            actual_file_count += 1
            current_time = currentDatetimeInDash


            log_file.write("================================================================\n")
            log_file.write(f"({actual_file_count}).{record["url"]}\n")
            log_file.write(f"Issuer:{record['company_name']}\n")
            log_file.write(f"Type :{normalized_form_type}\n")
            log_file.write(f"Total no of files merged : {len(document_links)}\n")
            log_file.write(f"Actual file count : {actual_file_count}\n")
            log_file.write(f"Time  : {current_time}\n")

        wb.save(so_excel_path)
        logger(f"Excel file created successfully: {excel_filename}", Logger.SUCCESS)

        # ----------------------------
        # Save Back to JSON
        # ----------------------------
        with open(info_json_path, "w", encoding="utf-8") as info_f:
            json.dump(info_json_data, info_f, indent=4)

        logger("Records successfully appended to info.json", Logger.SUCCESS)

    validation_status = check_missing_file_validation(run_folder)

    if not validation_status:
        logger("Missing PDF validation failed", Logger.ERROR)
    else:
        logger("PDF validation successful", Logger.SUCCESS)

    # "=============================================== SFTP TRANSFER ========================================================="
    upload_to_sftp(run_folder)

    # -----------------------------------
    # If run == 1 → create 1A and copy logs
    # -----------------------------------
    if str(current_run) == "1":
        source_folder = os.path.join(base_output_dir, "1")
        target_folder = os.path.join(base_output_dir, "1A")

        # Create 1A folder if not exists
        if not os.path.exists(target_folder):
            os.makedirs(target_folder)

        # Files to copy
        files_to_copy = ["scriptlog.txt", "scriptlog1.txt"]

        for file_name in files_to_copy:
            src_file = os.path.join(source_folder, file_name)
            dst_file = os.path.join(target_folder, file_name)

            if os.path.exists(src_file):
                shutil.copy2(src_file, dst_file)
            else:
                logger(f"File not found: {src_file}", Logger.WARNING)
        SO_Validator.validator_check_11(
            todayDate,
            target_folder,
            currentDatetimeInDash,
            forms
        )

        upload_to_sftp(target_folder)

        logger("The 1A/11 folder Created Successfully", Logger.SUCCESS)

    # Overwrite runcount.txt with current run
    with open(runcount_file, "w", encoding="utf-8") as f:
        f.write(f"{str(current_run)}")

    send_mail()
    print("--------SO Convertion Completed -------------")
    
except Exception as e:
    logger(f"Final Exception: {e}", Logger.ERROR)
    postmark.emails.send(
        From='XYZ@gmail.com',
        To='XXX@gmail.com',
        Subject='SEC_SO - LC [server_name] || Success!',
        HtmlBody=f'There are some issues were encountered. Check Log file. - SO : {current_run}.',
        Attachments=[log_file_name]
    )
finally:
    # ========================
    # DRIVER CLEANUP
    # ========================
    logger("========================================================================================================", Logger.SUCCESS)
    driver.quit()
