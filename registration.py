from pathlib import Path
from datetime import datetime
import subprocess
import sys

import requests
from openpyxl import load_workbook


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

RAW_FOLDER = BASE_DIR / "raw_wjx_data"
PARTICIPANTS_FOLDER = BASE_DIR / "participants_list"
SIGNUP_FOLDER = BASE_DIR / "signup_sheet"
HISTORY_FOLDER = BASE_DIR / "record_history"

RECORD_FILE = BASE_DIR / "record.xlsx"
LOG_FILE = BASE_DIR / "log.txt"

# WJX columns (1-based)
REGISTRATION_TIME_COLUMN = 2
NICKNAME_COLUMN = 7
STUDENT_ID_COLUMN = 8
REMARKS_COLUMN = 9

MINIMUM_COLUMNS = 9


# ============================================================
# BASIC UTILITIES
# ============================================================

def log(message=""):
    """
    Print a message to the console and append it to log.txt.

    Every output line receives its own timestamp in the format:
        [yyyy/mm/dd hh:mm:ss]: log content
    """

    timestamp = datetime.now().strftime(
        "%Y/%m/%d %H:%M:%S"
    )

    # Handle multi-line messages line by line so every output line
    # receives its own timestamp.
    lines = str(message).splitlines()

    if not lines:
        lines = [""]

    with LOG_FILE.open(
        "a",
        encoding="utf-8"
    ) as log_file:

        for line in lines:

            formatted = (
                f"[{timestamp}]: {line}"
            )

            print(formatted)
            log_file.write(
                formatted + "\n"
            )


def error(message):
    log()
    log("=" * 60)
    log("ERROR")
    log("=" * 60)
    log(message)
    log()
    sys.exit(1)
def normalize_student_id(value):
    """
    Convert a student ID into a consistent string.
    """

    if value is None:
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value).strip()


def parse_registration_time(value):
    """
    Convert WJX registration time into datetime.
    """

    if isinstance(value, datetime):
        return value

    if value is None:
        return None

    value = str(value).strip()

    formats = [
        "%Y/%m/%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass

    return None


# ============================================================
# DIRECTORY SETUP
# ============================================================

def create_directories():

    RAW_FOLDER.mkdir(
        parents=True,
        exist_ok=True
    )

    PARTICIPANTS_FOLDER.mkdir(
        parents=True,
        exist_ok=True
    )

    SIGNUP_FOLDER.mkdir(
        parents=True,
        exist_ok=True
    )

    HISTORY_FOLDER.mkdir(
        parents=True,
        exist_ok=True
    )


# ============================================================
# SERIAL NUMBER
# ============================================================

def get_next_serial_number(date_string):

    existing_numbers = []

    pattern = f"raw_wjx_data_{date_string}_*.xlsx"

    for file in RAW_FOLDER.glob(pattern):

        try:
            number = int(
                file.stem.split("_")[-1]
            )

            if number >= 1:
                existing_numbers.append(number)

        except ValueError:
            pass

    if not existing_numbers:
        return 1

    return max(existing_numbers) + 1


# ============================================================
# FILE SERIAL PARSING
# ============================================================

def get_filename_date_and_serial(file, prefix):
    """
    Extract date and serial number from filenames such as:

        signup_sheet_20260920_003.xlsx
        record_history_20260920_003.xlsx
    """

    stem = file.stem

    expected_start = prefix + "_"

    if not stem.startswith(expected_start):
        return None

    remainder = stem[len(expected_start):]

    parts = remainder.split("_")

    if len(parts) != 2:
        return None

    date_string = parts[0]
    serial_string = parts[1]

    if len(date_string) != 8:
        return None

    try:
        serial = int(serial_string)
    except ValueError:
        return None

    return (
        date_string,
        serial
    )


# ============================================================
# RECORD HISTORY CHECK
# ============================================================

def get_latest_signup_sheet():

    files = list(
        SIGNUP_FOLDER.glob(
            "signup_sheet_*.xlsx"
        )
    )

    parsed = []

    for file in files:

        result = get_filename_date_and_serial(
            file,
            "signup_sheet"
        )

        if result is not None:
            parsed.append(
                (result, file)
            )

    if not parsed:
        return None

    parsed.sort(
        key=lambda x: x[0]
    )

    return parsed[-1][1]


def history_exists_for_signup_sheet(
    signup_file
):
    """
    Check whether the corresponding record_history file
    exists for the latest signup sheet.
    """

    result = get_filename_date_and_serial(
        signup_file,
        "signup_sheet"
    )

    if result is None:
        return False

    date_string, serial = result

    history_file = (
        HISTORY_FOLDER
        / f"record_history_{date_string}_{serial:03d}.xlsx"
    )

    return history_file.exists()


def ensure_record_is_up_to_date():
    """
    Before starting a new registration:

    If there is a previous signup sheet without a corresponding
    record_history file, ask whether record.py should be run.

    Y = automatically run record.py
    n = terminate
    """

    latest_signup = get_latest_signup_sheet()

    # No previous signup sheet means this is the first event.
    if latest_signup is None:
        return

    if history_exists_for_signup_sheet(
        latest_signup
    ):
        return

    log()
    log("=" * 60)
    log("WARNING")
    log("=" * 60)

    log(
        "\nThe latest signup sheet does not have a corresponding "
        "record history."
    )

    log()
    log(
        f"Latest signup sheet:\n"
        f"  {latest_signup.name}"
    )

    log()
    log(
        "The attendance record may therefore not be up to date."
    )

    log()
    log(
        "Enter Y to automatically run record.py."
    )

    log(
        "Enter n to terminate and run record.py manually."
    )

    log()

    choice = input(
        "Your choice [Y/n]: "
    ).strip()

    # Only uppercase Y triggers automatic execution.
    if choice == "Y":

        log()
        log("Running record.py automatically...")
        log()

        record_script = (
            BASE_DIR / "record.py"
        )

        if not record_script.exists():

            error(
                "record.py was not found.\n\n"
                f"Expected location:\n"
                f"  {record_script}"
            )

        try:

            result = subprocess.run(
                [
                    sys.executable,
                    str(record_script)
                ],
                cwd=BASE_DIR
            )

        except Exception as e:

            error(
                "Could not run record.py.\n\n"
                f"{e}"
            )

        if result.returncode != 0:

            error(
                "record.py did not finish successfully.\n\n"
                "The registration process has been stopped."
            )

        # Verify that record.py actually produced the required
        # history file.
        if not history_exists_for_signup_sheet(
            latest_signup
        ):

            error(
                "record.py finished, but the required "
                "record_history file was not found.\n\n"
                "The registration process has been stopped."
            )

        log()
        log(
            "Attendance record successfully updated."
        )

    else:

        log()
        log(
            "Please run record.py manually before starting "
            "a new registration."
        )

        log()
        sys.exit(0)


# ============================================================
# DOWNLOAD
# ============================================================

def download_file(
    url,
    output_path
):

    log()
    log("Downloading WJX file...")

    try:

        response = requests.get(
            url,
            stream=True,
            timeout=60
        )

        response.raise_for_status()

    except requests.RequestException as e:

        error(
            "Failed to download the WJX file.\n\n"
            f"{e}"
        )

    try:

        with open(
            output_path,
            "wb"
        ) as file:

            for chunk in response.iter_content(
                chunk_size=1024 * 1024
            ):

                if chunk:
                    file.write(chunk)

    except OSError as e:

        error(
            "Could not save the downloaded file.\n\n"
            f"{e}"
        )

    log(
        f"Downloaded successfully:\n"
        f"  {output_path}"
    )


# ============================================================
# VALIDATE WJX
# ============================================================

def validate_wjx_file(
    input_file
):

    log()
    log("Validating downloaded Excel file...")

    try:

        workbook = load_workbook(
            input_file,
            read_only=True,
            data_only=True
        )

    except Exception as e:

        error(
            "The downloaded file could not be opened as "
            "an Excel file.\n\n"
            f"{e}"
        )

    worksheet = workbook.active

    if worksheet.max_column < MINIMUM_COLUMNS:

        workbook.close()

        error(
            f"The downloaded file only has "
            f"{worksheet.max_column} columns.\n\n"
            f"This program expects at least "
            f"{MINIMUM_COLUMNS} columns."
        )

    if worksheet.max_row < 1:

        workbook.close()

        error(
            "The downloaded Excel file is empty."
        )

    workbook.close()

    log("Excel file validation passed.")


# ============================================================
# RECORD FILE
# ============================================================

def validate_record_file():

    if not RECORD_FILE.exists():

        error(
            f"record.xlsx was not found.\n\n"
            f"Expected location:\n"
            f"  {RECORD_FILE}"
        )

    try:

        workbook = load_workbook(
            RECORD_FILE,
            read_only=True,
            data_only=True
        )

    except Exception as e:

        error(
            "record.xlsx could not be opened.\n\n"
            f"{e}"
        )

    worksheet = workbook.active

    # record.xlsx is allowed to be completely empty before the first event.
    # In that case, there are simply no previous attendance records and every
    # student will be treated as having an attendance count of 0.
    is_empty = (
        worksheet.max_row == 1
        and worksheet.max_column == 1
        and worksheet["A1"].value is None
    )

    if not is_empty and worksheet.max_column < 2:

        workbook.close()

        error(
            "record.xlsx must contain at least two columns:\n\n"
            "Column 1 = student ID\n"
            "Column 2 = attendance count"
        )

    workbook.close()


def load_record():

    workbook = load_workbook(
        RECORD_FILE,
        data_only=True
    )

    worksheet = workbook.active

    records = {}

    for row in worksheet.iter_rows(
        min_row=2,
        values_only=True
    ):

        if not row:
            continue

        student_id = normalize_student_id(
            row[0]
        )

        if not student_id:
            continue

        count = row[1]

        if count is None:
            count = 0

        try:

            count = int(count)

        except (ValueError, TypeError):

            workbook.close()

            error(
                f"Invalid attendance count for "
                f"student {student_id}."
            )

        records[student_id] = count

    workbook.close()

    return records


# ============================================================
# LOAD WJX REGISTRATIONS
# ============================================================

def load_registrations(
    input_file
):

    workbook = load_workbook(
        input_file,
        data_only=True
    )

    worksheet = workbook.active

    registrations_by_id = {}

    total_rows = 0
    duplicate_count = 0

    for row in worksheet.iter_rows(
        min_row=2,
        values_only=True
    ):

        total_rows += 1

        if len(row) < STUDENT_ID_COLUMN:
            continue

        student_id = normalize_student_id(
            row[STUDENT_ID_COLUMN - 1]
        )

        if not student_id:
            continue

        nickname = row[
            NICKNAME_COLUMN - 1
        ]

        registration_time = parse_registration_time(
            row[
                REGISTRATION_TIME_COLUMN - 1
            ]
        )

        remarks = row[
            REMARKS_COLUMN - 1
        ]

        registration = {
            "student_id": student_id,
            "nickname": nickname,
            "registration_time": registration_time,
            "remarks": remarks,
            "original_row": list(row)
        }

        # ----------------------------------------------------
        # DEDUPLICATION
        # ----------------------------------------------------

        if student_id not in registrations_by_id:

            registrations_by_id[
                student_id
            ] = registration

        else:

            duplicate_count += 1

            existing = registrations_by_id[
                student_id
            ]

            old_time = existing[
                "registration_time"
            ]

            new_time = registration[
                "registration_time"
            ]

            # Keep earliest registration.
            if (
                old_time is None
                and new_time is not None
            ):

                registrations_by_id[
                    student_id
                ] = registration

            elif (
                old_time is not None
                and new_time is not None
                and new_time < old_time
            ):

                registrations_by_id[
                    student_id
                ] = registration

    workbook.close()

    registrations = list(
        registrations_by_id.values()
    )

    log()
    log(
        f"Rows in WJX file: "
        f"{total_rows}"
    )

    log(
        f"Unique student IDs: "
        f"{len(registrations)}"
    )

    log(
        f"Duplicate registrations removed: "
        f"{duplicate_count}"
    )

    return registrations


# ============================================================
# SELECT PARTICIPANTS
# ============================================================

def select_participants(
    registrations,
    records,
    limit
):

    def sort_key(student):

        student_id = student[
            "student_id"
        ]

        attendance = records.get(
            student_id,
            0
        )

        registration_time = (
            student["registration_time"]
            if student["registration_time"] is not None
            else datetime.max
        )

        return (
            attendance,
            registration_time
        )

    sorted_students = sorted(
        registrations,
        key=sort_key
    )

    return sorted_students[:limit]


# ============================================================
# PARTICIPANT LIST
# ============================================================

def save_participant_list(
    output_file,
    selected_participants
):
    """
    Create:

        participants_list_yyyymmdd_no.xlsx

    Column 1 = nickname
    Column 2 = student ID
    Column 3 = remarks

    The remarks are taken from column 9 of the downloaded WJX file.
    """

    from openpyxl import Workbook

    workbook = Workbook()
    worksheet = workbook.active

    worksheet.title = "Participants"

    # Header
    worksheet.cell(
        row=1,
        column=1
    ).value = "Nickname"

    worksheet.cell(
        row=1,
        column=2
    ).value = "Student ID"

    worksheet.cell(
        row=1,
        column=3
    ).value = "Remarks"

    # Students
    for row_number, participant in enumerate(
        selected_participants,
        start=2
    ):

        worksheet.cell(
            row=row_number,
            column=1
        ).value = participant["nickname"]

        worksheet.cell(
            row=row_number,
            column=2
        ).value = participant["student_id"]

        worksheet.cell(
            row=row_number,
            column=3
        ).value = participant["remarks"]

    workbook.save(
        output_file
    )

    workbook.close()


# ============================================================
# SIGNUP SHEET
# ============================================================

def save_signup_sheet(
    output_file,
    selected_participants
):
    """
    Create:

        signup_sheet_yyyymmdd_no.xlsx

    Column 1 = nickname
    Column 2 = student ID
    Column 3 = attendance mark

    Blank = attended
    1     = absent
    """

    workbook = load_workbook(
        RECORD_FILE,
        data_only=True
    )

    # Create a completely new workbook instead of modifying
    # record.xlsx.
    workbook.close()

    from openpyxl import Workbook

    signup_workbook = Workbook()

    worksheet = signup_workbook.active

    worksheet.title = "Signup"

    # Header
    worksheet.cell(
        row=1,
        column=1
    ).value = "Nickname"

    worksheet.cell(
        row=1,
        column=2
    ).value = "Student ID"

    worksheet.cell(
        row=1,
        column=3
    ).value = "Attendance"

    # Students
    for row_number, participant in enumerate(
        selected_participants,
        start=2
    ):

        worksheet.cell(
            row=row_number,
            column=1
        ).value = participant[
            "nickname"
        ]

        worksheet.cell(
            row=row_number,
            column=2
        ).value = participant[
            "student_id"
        ]

        # Leave attendance blank.
        worksheet.cell(
            row=row_number,
            column=3
        ).value = None

    signup_workbook.save(
        output_file
    )

    signup_workbook.close()


# ============================================================
# MAIN
# ============================================================

def main():

    log()
    log("=" * 60)
    log("WJX REGISTRATION PROCESSOR")
    log("=" * 60)

    # --------------------------------------------------------
    # Create directories
    # --------------------------------------------------------

    create_directories()

    # --------------------------------------------------------
    # Check whether previous attendance has been recorded
    # --------------------------------------------------------

    ensure_record_is_up_to_date()

    # --------------------------------------------------------
    # Validate record.xlsx
    # --------------------------------------------------------

    log()
    log("Checking record.xlsx...")

    validate_record_file()

    # --------------------------------------------------------
    # URL
    # --------------------------------------------------------

    log()

    url = input(
        "Enter the WJX download URL:\n> "
    ).strip()

    if not url:

        error(
            "The download URL cannot be empty."
        )

    # --------------------------------------------------------
    # Participant limit
    # --------------------------------------------------------

    log()

    limit_input = input(
        "Enter the maximum number of participants:\n> "
    ).strip()

    try:

        participant_limit = int(
            limit_input
        )

        if participant_limit <= 0:
            raise ValueError

    except ValueError:

        error(
            "The participant limit must be "
            "a positive integer."
        )

    # --------------------------------------------------------
    # Date / serial
    # --------------------------------------------------------

    today = datetime.now()

    date_string = today.strftime(
        "%Y%m%d"
    )

    serial_number = get_next_serial_number(
        date_string
    )

    serial_string = f"{serial_number:03d}"

    log()
    log(
        f"Date:              {date_string}"
    )

    log(
        f"Serial number:     {serial_string}"
    )

    log(
        f"Participant limit: {participant_limit}"
    )

    # --------------------------------------------------------
    # Filenames
    # --------------------------------------------------------

    raw_file = (
        RAW_FOLDER
        / f"raw_wjx_data_"
        f"{date_string}_"
        f"{serial_string}.xlsx"
    )

    participant_file = (
        PARTICIPANTS_FOLDER
        / f"participants_list_"
        f"{date_string}_"
        f"{serial_string}.xlsx"
    )

    signup_file = (
        SIGNUP_FOLDER
        / f"signup_sheet_"
        f"{date_string}_"
        f"{serial_string}.xlsx"
    )

    # --------------------------------------------------------
    # Download
    # --------------------------------------------------------

    download_file(
        url,
        raw_file
    )

    # --------------------------------------------------------
    # Validate
    # --------------------------------------------------------

    validate_wjx_file(
        raw_file
    )

    # --------------------------------------------------------
    # Load attendance record
    # --------------------------------------------------------

    log()
    log("Reading attendance records...")

    records = load_record()

    log(
        f"Students currently in record.xlsx: "
        f"{len(records)}"
    )

    # --------------------------------------------------------
    # Read registrations
    # --------------------------------------------------------

    log()
    log("Reading WJX registrations...")

    registrations = load_registrations(
        raw_file
    )

    if not registrations:

        error(
            "No valid student registrations were found."
        )

    # --------------------------------------------------------
    # Select
    # --------------------------------------------------------

    log()
    log("Selecting participants...")

    selected_participants = select_participants(
        registrations,
        records,
        participant_limit
    )

    log()
    log(
        f"Selected "
        f"{len(selected_participants)} "
        f"participants."
    )

    # --------------------------------------------------------
    # Display
    # --------------------------------------------------------

    log()
    log("Selected students:")

    for index, participant in enumerate(
        selected_participants,
        start=1
    ):

        student_id = participant[
            "student_id"
        ]

        nickname = participant[
            "nickname"
        ]

        previous_count = records.get(
            student_id,
            0
        )

        log(
            f"  {index:03d}. "
            f"{student_id} | "
            f"{nickname} | "
            f"previous attendance: "
            f"{previous_count}"
        )

    # --------------------------------------------------------
    # Participant list
    # --------------------------------------------------------

    log()
    log("Creating participant list...")

    save_participant_list(
        participant_file,
        selected_participants
    )

    log(
        f"Saved:\n"
        f"  {participant_file}"
    )

    # --------------------------------------------------------
    # Signup sheet
    # --------------------------------------------------------

    log()
    log("Creating signup sheet...")

    save_signup_sheet(
        signup_file,
        selected_participants
    )

    log(
        f"Saved:\n"
        f"  {signup_file}"
    )

    # --------------------------------------------------------
    # Final
    # --------------------------------------------------------

    log()
    log("=" * 60)
    log("REGISTRATION COMPLETED")
    log("=" * 60)

    log()
    log(
        "After the event, open the signup sheet and:"
    )

    log(
        "  Leave column 3 blank for students who attended "
        "or were absent without a valid reason."
    )

    log(
        "  Enter 1 in column 3 for students who were absent "
        "with a valid reason."
    )

    log()
    log(
        "Then run record.bat before the next registration."
    )

    log()
    log("=" * 60)


if __name__ == "__main__":
    main()