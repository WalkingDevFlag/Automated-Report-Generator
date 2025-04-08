# config.py
"""
Configuration module for the Automated Report Generator.

Loads environment variables, defines constants for Google API scopes,
spreadsheet details, file paths, email settings, and crucially,
maps Google Sheet column headers to constants for reliable data access
across different report types.
"""

import os
import sys
from dotenv import load_dotenv

# --- Environment Variable Loading ---
# Load environment variables from .env file
# Ensure you have a .env file in the project root with:
# SPREADSHEET_ID=YOUR_SHEET_ID
# EMAIL_SENDER=your_email@example.com
# EMAIL_PASSWORD=your_email_app_password_or_regular_password
load_dotenv()
print("[INFO] Loading configuration from environment variables...")

# --- Sensitive Configuration (Loaded from .env) ---
EMAIL_SENDER = os.getenv('EMAIL_SENDER')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')

# --- Validation: Check if essential environment variables are set ---
missing_vars = []
if not EMAIL_SENDER: missing_vars.append("EMAIL_SENDER")
if not EMAIL_PASSWORD: missing_vars.append("EMAIL_PASSWORD")
if not SPREADSHEET_ID: missing_vars.append("SPREADSHEET_ID")

if missing_vars:
    print(f"[ERROR] Missing required environment variables in .env file: {', '.join(missing_vars)}")
    sys.exit(f"Exiting: Configuration incomplete. Please define {', '.join(missing_vars)} in your .env file.")

# --- Google API Configuration ---
# Scopes define the level of access requested from Google APIs
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly', # Read data from sheets
    'https://www.googleapis.com/auth/drive.readonly'        # Read files from drive (including exports)
]

# --- Spreadsheet Configuration ---
# Target sheet name and data range within the spreadsheet.
# Example: If your data is in 'Sheet1' from column A to column M, use 'Sheet1!A:M'.
# Ensure the range covers ALL columns needed by ANY report type.
SHEET_NAME = 'Form Responses 1' # <<<--- ADJUST SHEET NAME IF DIFFERENT
LAST_COLUMN = 'AZ'              # <<<--- ADJUST LAST COLUMN LETTER NEEDED
RANGE_NAME = f'{SHEET_NAME}!A:{LAST_COLUMN}'

# --- File/Folder Paths ---
# Assumes this config.py file is in the root directory of your project.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(BASE_DIR, 'credentials.json') # Google OAuth 2.0 Client Secrets file
TOKEN_FILE = os.path.join(BASE_DIR, 'token.json')             # Stores user's access and refresh tokens
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'generated_reports')    # Where successfully generated reports are saved
TEMP_FOLDER = os.path.join(BASE_DIR, 'temp_files')             # Temporary storage for downloads, conversions, etc.
TEMPLATE_FOLDER = os.path.join(BASE_DIR, 'template')           # Contains the .docx template files

# --- Email Server Configuration ---
SMTP_SERVER = 'smtp.gmail.com' # Default for Gmail
SMTP_PORT = 465                # Default for Gmail with SSL (use 587 for TLS)
# Choose port based on your email provider's requirements (465=SSL, 587=TLS)
EMAIL_USE_SSL = (SMTP_PORT == 465) # Helper flag based on port

# --- Google Sheet Column Headers ---
# These MUST EXACTLY MATCH the headers in the first row of your Google Sheet (defined by RANGE_NAME).
# Using constants prevents errors from typos when accessing data.

# --- Common Headers (Present in most/all responses) ---
HEADER_TIMESTAMP = 'Timestamp'                              # Timestamp of form submission
HEADER_NAME = 'Name'                                        # Name of the submitter
HEADER_EMAIL_RECIPIENT = 'Email Address to receive the report' # Email address for sending the report
HEADER_TEMPLATE_CHOICE = 'What report do you want?'         # User's selected report type

# --- Event Report Specific Headers ---
HEADER_EVENT_INTRO = 'Introduction of the Event'
HEADER_EVENT_OBJECTIVE = 'Objective of the Event (bullet points or about 50 words)'
HEADER_EVENT_TITLE = 'Event Name / Title'
HEADER_EVENT_LOCATION = 'Location/Place'
HEADER_EVENT_DETAILS = 'Content/Details'
HEADER_EVENT_START_DATE = 'Start Date'
HEADER_EVENT_END_DATE = 'End Date'
HEADER_EVENT_IMAGE_1 = 'Event Brochure'                     # Google Drive URL (Image)
HEADER_EVENT_IMAGE_2 = 'Event Images'                       # Google Drive URL (Image)

# --- Remedial Class Report Specific Headers ---
HEADER_REMEDIAL_SCHOOL = 'Name of the School'
HEADER_REMEDIAL_DEPARTMENT = 'Name of the Department'
HEADER_REMEDIAL_PROGRAM = 'Name of the Program'
HEADER_REMEDIAL_COURSE = 'Name of the Course'
HEADER_REMEDIAL_COURSE_CODE = 'Course Code'
HEADER_REMEDIAL_COURSE_CREDIT = 'Course Credit'
HEADER_REMEDIAL_TOTAL_HOURS = 'Total Hours of Course as Per DMS'
HEADER_REMEDIAL_CLASSES_TAKEN = 'Total number of remedial classes taken'
HEADER_REMEDIAL_FACULTY = 'Name of Faculty Member'
HEADER_REMEDIAL_SESSION = 'Session'
HEADER_REMEDIAL_SEMESTER = 'Semester'
HEADER_REMEDIAL_CRITERIA = 'Criteria for Identification of slow learners for remedial class'
HEADER_REMEDIAL_STUDENT_LIST_CSV = 'Upload a list of identified students for Remedial classes and Performance analysis' # Google Drive URL (CSV/Sheet)
HEADER_REMEDIAL_TIMETABLE_CSV = 'Upload Remedial Classes details (Timetable)'      # Google Drive URL (CSV/Sheet)
HEADER_REMEDIAL_ATTENDANCE_CSV = 'Upload Attendance of Students'           # Google Drive URL (CSV/Sheet)

# --- Curriculum Map Report Specific Headers ---
HEADER_CM_FACULTY = 'Faculty'
HEADER_CM_SCHOOL = 'Name of School'
HEADER_CM_DEPARTMENT = 'Name of Department'
HEADER_CM_EVENT_NAME = 'Name of the Event'
HEADER_CM_NATURE = 'Nature of Event'
HEADER_CM_DATE = 'Date of Event'
HEADER_CM_INTRODUCTION = 'Introduction of Event'
HEADER_CM_OBJECTIVE = 'Objective of the Event'
HEADER_CM_BENEFICIARIES = 'Beneficiaries of the Event'
HEADER_CM_GUESTS_CSV = 'Details of the Guests'              # Google Drive URL (CSV/Sheet)
HEADER_CM_DESCRIPTION = 'Brief Description of the Event'
HEADER_CM_PROGRAM_CSV = 'Program Scheme'                    # Google Drive URL (CSV/Sheet)
HEADER_CM_LECTURE_PLAN_CSV = 'Lecture Plan'                 # Google Drive URL (CSV/Sheet)
HEADER_CM_PHOTOS = 'Geo-tagged Photographs'                 # Google Drive URL (Image)
HEADER_CM_BROCHURE = 'Brochure or Creative of the Event'    # Google Drive URL (Image)
HEADER_CM_SCHEDULE_CSV = 'Schedule of the Event'            # Google Drive URL (CSV/Sheet)
HEADER_CM_ATTENDANCE_CSV = 'Attendance of the Event'        # Google Drive URL (CSV/Sheet)
HEADER_CM_NEWS = 'News Publication'                         # Google Drive URL (Image/PDF/Doc)
HEADER_CM_FEEDBACK = 'Feedback of the Event'
HEADER_CM_LINK = 'Link of MUJ website'
HEADER_CM_CLUB_LOGO = 'Logo Of Club'                        # Google Drive URL (Image)

# --- Summarizer Report Specific Headers ---
HEADER_SUMMARY_UPLOAD_FILE = 'Upload Report'                # Google Drive URL (PDF/DOCX)

# --- Image/Asset Handling Defaults ---
IMAGE_WIDTH_INCHES = 5.0    # Default width for Event Report images
LOGO_WIDTH_INCHES = 1.0     # Default width for Curriculum Map logo
# Note: Summarizer image width is defined within report_summarizer.py

print("[SUCCESS] Configuration loaded successfully.")