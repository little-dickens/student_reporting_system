import os
from datetime import datetime
from dateutil import parser
import sqlite3
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DB_FILE_PATH = PROJECT_ROOT / "database" / "master.db"

def format_date(date_str, display_format=None, out=False):
    # from gui to db -> YYYY-MM-DD
    if out:
        if date_str:
            datetime_obj = parser.parse(date_str)
            # Convert date to default datestring for sqlite: YYYY-MM-DD
            formatted_date = datetime_obj.strftime("%Y-%m-%d")
            return formatted_date
        else:
            return ""
    # from db to gui -> MM-DD-YYYY
    else:
        if date_str:
            datetime_obj = parser.parse(date_str)
            # Remove leading zeros -> MM-DD-YYYY
            if os.name != 'nt':
                if display_format:
                    formatted_date = datetime_obj.strftime(display_format)
                else:    
                    formatted_date = datetime_obj.strftime("%-m-%-d-%Y")
            else:
                if display_format:
                    formatted_date = datetime_obj.strftime(display_format)
                else:    
                    formatted_date = datetime_obj.strftime("%#m-%#d-%Y")
            return formatted_date
        else:
            return ""        

def format_mobile(mobile_str, out=False):
    # gui -> db 1234567891 (only numeric)
    if out:
        if len(mobile_str) == 0:
            return ""
        else:
            num_str = ""
            for char in mobile_str:
                if char.isnumeric():
                    num_str += char
            return int(num_str)
    # db -> gui: (123) 456-7891
    else:
        # if len(mobile_str) == 0 or mobile_str is None:
        if not mobile_str:
            return mobile_str
        else:    
            num_str = ""
            for char in mobile_str:
                if char.isdigit():
                    num_str += char
            formatted_num = f"({num_str[:3]}) {num_str[3:6]}-{num_str[6:]}"
            return formatted_num
        
def get_test_png_path():
    # Get the absolute path of the current file
    current_path = os.path.realpath(__file__)
    # Split the current path into components based on the platform-specific separator
    path_parts = current_path.split(os.path.sep)
    # Find the index of 'student_reporting_system' in the path
    try:
        ind = path_parts.index('student_reporting_system')
    except ValueError:
        raise RuntimeError("'student_reporting_system' not found in the current path")

    # Construct the file stem using path segments up to 'student_reporting_system'
    file_stem = os.path.join(*path_parts[:ind + 1])

    # Append the database path
    pngs_path = os.path.join(file_stem, 'design')
    pngs_path = os.path.join(pngs_path, 'pngs')

    # Ensure the path starts with C:\ on Windows
    if os.name == 'nt':  # 'nt' is the name for Windows
        if pngs_path.startswith('C:'):
            pngs_path = pngs_path.replace('C:', 'C:\\', 1)
    elif os.name == 'posix':  # 'posix' is the name for Unix-like systems (Linux, macOS)
        if not pngs_path.startswith('/'):
            pngs_path = f'/{pngs_path.lstrip("/")}'
    return pngs_path

def format_date_range(date_range):
    """
    Converts date_range to a tuple of strings in the format ('YYYY-MM-DD', 'YYYY-MM-DD').

    Parameters:
    - date_range (tuple): A tuple containing two dates, either as datetime objects or strings.

    Returns:
    - tuple: A tuple of strings in the format ('YYYY-MM-DD', 'YYYY-MM-DD').
    """
    # Default date range if input is invalid
    default_start = '2024-07-01'
    default_end = '2025-06-30'
    
    try:
        # Unpack start and end dates from date_range
        start, end = date_range

        # Convert datetime objects to strings
        if isinstance(start, datetime):
            start = start.strftime('%Y-%m-%d')
        if isinstance(end, datetime):
            end = end.strftime('%Y-%m-%d')

        # Check if start and end are now strings in the desired format
        datetime.strptime(start, '%Y-%m-%d')  # Validate format
        datetime.strptime(end, '%Y-%m-%d')    # Validate format
        
        return (start, end)
    except (ValueError, TypeError):
        # If there's an error with format, return a default range
        print("Invalid date range format. Using default range.")
        return (default_start, default_end)

def get_site_names(ignore_empty=False):
    
    conn = sqlite3.connect(DB_FILE_PATH)
    cursor = conn.cursor()

    if ignore_empty:

        # Fetch sites with non-zero attendance
        query = '''
            SELECT DISTINCT
                site
            FROM
                Attendance
        '''
    else:
        # Fetch the class record to get the full notes text
        query = '''
                SELECT
                    site_name
                FROM
                    Sites
                ORDER BY
                    site_name ASC
        '''
    cursor.execute(query)
    site_names = [value for tup in cursor.fetchall() for value in tup]
    conn.close()
    return site_names

def get_test_dates_by_site(site):
    
    conn = sqlite3.connect(DB_FILE_PATH)
    cursor = conn.cursor()

     # Fetch the class record to get the full notes text
    query = '''
        SELECT DISTINCT date
            FROM Testing
            WHERE site = ?
            ORDER BY date DESC
    '''
    cursor.execute(query, (site,))
    test_dates = [value for tup in cursor.fetchall() for value in tup]
    conn.close()
    return test_dates