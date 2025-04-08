# main.py
"""
Main orchestration script for the Automated Report Generator.

Handles:
- Initialization and setup (folder creation).
- Google API Authentication (OAuth 2.0).
- Fetching data from the configured Google Sheet.
- Iterating through sheet rows (responses).
- Identifying the requested report type.
- Downloading required assets (images, CSVs, PDFs/DOCX) from Google Drive.
- Calling the appropriate report generation module (event, remedial, curriculum, summarizer).
- Saving the generated DOCX report.
- Emailing the report to the specified recipient.
- Cleaning up temporary files for successfully processed and emailed reports.
"""
# --- Ensure __future__ import is at the very top ---
from __future__ import annotations
import os
import io
import re
import sys # For sys.exit
from datetime import datetime
import shutil # For potential future use (e.g., removing temp folder - currently disabled)
import traceback # For detailed error printing
import time # For timing operations



# --- Google Cloud Libraries ---
# Check for availability and provide guidance if missing
try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    from googleapiclient.http import MediaIoBaseDownload
    GOOGLE_LIBS_AVAILABLE = True
except ImportError:
    print("[ERROR] Required Google libraries ('google-auth-oauthlib', 'google-api-python-client') not found.")
    print("  >>> Please install them: pip install google-auth-oauthlib google-api-python-client")
    GOOGLE_LIBS_AVAILABLE = False
    sys.exit("Exiting: Core Google libraries are required.")

# --- Image Conversion Library (Optional but Recommended) ---
try:
    from PIL import Image, UnidentifiedImageError
    PIL_INSTALLED = True
except ImportError:
    PIL_INSTALLED = False
    print("[WARNING] Optional library 'Pillow' not found (pip install Pillow).")
    print("  >>> Image conversion (e.g., WEBP to PNG) and advanced image handling will be disabled.")

# --- Project Modules ---
# Check for availability and provide guidance if missing
try:
    import config          # Load configuration constants and .env variables
    import email_sender    # Handles sending emails
    # Report-specific modules
    import event_report
    import remedial_report
    import curriculum_map_report
    import report_summarizer
    MODULES_AVAILABLE = True
except ImportError as import_err:
     print(f"[ERROR] Failed to import one or more project modules: {import_err}")
     print("  >>> Ensure all .py files (config, email_sender, report modules) are present in the directory.")
     MODULES_AVAILABLE = False
     sys.exit("Exiting: Core project modules are missing.")


# ===========================================================
# --- Helper Functions ---
# ===========================================================

def get_timestamp() -> str:
    """Returns the current timestamp in a standard format for logging."""
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def authenticate() -> Credentials | None:
    """
    Handles Google Authentication using OAuth 2.0 credentials.

    Loads existing tokens from token.json, refreshes if expired,
    or initiates the OAuth flow for the user if no valid token exists.

    Returns:
        Credentials | None: Valid Google API credentials object, or None on failure.
    """
    print(f"\n=== {get_timestamp()} AUTHENTICATION ===")
    creds = None
    token_path = config.TOKEN_FILE
    credentials_path = config.CREDENTIALS_FILE

    # --- Load credentials from token file if exists ---
    if os.path.exists(token_path):
        try:
            creds = Credentials.from_authorized_user_file(token_path, config.SCOPES)
            print(f"[INFO] Loaded credentials from {os.path.basename(token_path)}")
        except Exception as e:
            print(f"[WARNING] Error loading token file '{os.path.basename(token_path)}': {e}. Will attempt re-authentication.")
            try: os.remove(token_path) # Remove potentially corrupted token file
            except OSError: pass
            creds = None

    # --- Check if credentials are valid or need refresh ---
    if creds and creds.valid:
        print("[SUCCESS] Existing credentials are valid.")
        return creds
    elif creds and creds.expired and creds.refresh_token:
        print("[INFO] Credentials expired. Attempting to refresh token...")
        try:
            creds.refresh(Request())
            print("[SUCCESS] Token refreshed successfully.")
            # --- Save the refreshed credentials ---
            try:
                with open(token_path, 'w') as token_file:
                    token_file.write(creds.to_json())
                print(f"[INFO] Refreshed token saved to {os.path.basename(token_path)}")
            except Exception as e:
                print(f"[WARNING] Error saving refreshed token: {e}")
            return creds # Return refreshed credentials
        except Exception as e:
            print(f"[ERROR] Failed to refresh token: {e}")
            print("  >>> Will proceed to full re-authentication.")
            try: os.remove(token_path) # Remove invalid token file
            except OSError: pass
            creds = None # Force re-authentication

    # --- Initiate OAuth flow if no valid credentials ---
    if not creds:
        print("[INFO] No valid credentials found. Starting OAuth 2.0 flow...")
        if not os.path.exists(credentials_path):
             print(f"[ERROR] OAuth credentials file ('{os.path.basename(credentials_path)}') not found.")
             print("  >>> Please download your OAuth 2.0 Client ID credentials from Google Cloud Console and save as 'credentials.json'.")
             return None
        try:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, config.SCOPES)
            # run_local_server will open a browser tab for user authorization
            # Ensure redirect_uris in credentials.json includes "http://localhost:port"
            print("[ACTION] Please follow the prompts in your web browser to authorize access.")
            creds = flow.run_local_server(port=0) # Use port=0 for random available port
            print("[SUCCESS] OAuth flow completed.")
            # --- Save the new credentials ---
            try:
                with open(token_path, 'w') as token_file:
                    token_file.write(creds.to_json())
                print(f"[INFO] New token saved to {os.path.basename(token_path)}")
            except Exception as e:
                print(f"[WARNING] Error saving new token file: {e}")
            return creds
        except FileNotFoundError:
             # This case should be caught by the os.path.exists check above, but included for safety.
             print(f"[ERROR] Credentials file '{os.path.basename(credentials_path)}' not found during flow.")
             return None
        except Exception as e:
             print(f"[ERROR] An error occurred during the OAuth authentication flow: {e}")
             traceback.print_exc()
             return None

    # Should not be reachable if logic is correct, but return None as fallback
    return None

def get_sheet_data(creds: Credentials) -> tuple[list[str] | None, list[list[str]] | None]:
    """
    Fetches data (headers and rows) from the Google Sheet specified in config.

    Args:
        creds (Credentials): Valid Google API credentials object.

    Returns:
        tuple[list[str] | None, list[list[str]] | None]:
            - A list containing the header row strings, or None on failure.
            - A list of lists, where each inner list represents a data row, or None on failure.
              Returns an empty list for data rows if only the header exists.
    """
    print(f"\n=== {get_timestamp()} DATA FETCHING ===")
    print(f"[INFO] Fetching data from Sheet ID: {config.SPREADSHEET_ID}")
    print(f"           Range: {config.RANGE_NAME}")
    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().get(
            spreadsheetId=config.SPREADSHEET_ID,
            range=config.RANGE_NAME
        ).execute()

        values = result.get('values', [])

        if not values:
            print("[WARNING] No data found in the specified sheet/range.")
            return None, None # Indicate failure clearly

        headers = values[0]
        print(f"[INFO] Found {len(headers)} header columns.")
        # print(f"  [DEBUG] Headers: {headers}") # Uncomment for debugging

        # Check if only header row exists
        if len(values) < 2:
            print("[INFO] No data rows found (only header row).")
            return headers, [] # Return headers but empty data list

        data_rows = values[1:]
        print(f"[SUCCESS] Fetched {len(data_rows)} data row(s).")
        return headers, data_rows

    except HttpError as err:
        print(f"[ERROR] Google Sheets API HTTP error occurred: {err}")
        if err.resp.status == 403:
             print("  >>> Suggestion: Check API permissions for Sheets API and sharing settings for the spreadsheet.")
        elif err.resp.status == 404:
             print(f"  >>> Suggestion: Verify the SPREADSHEET_ID ('{config.SPREADSHEET_ID}') and RANGE_NAME ('{config.RANGE_NAME}') are correct.")
        traceback.print_exc()
        return None, None
    except Exception as e:
         print(f"[ERROR] An unexpected error occurred getting sheet data: {e}")
         traceback.print_exc()
         return None, None

def extract_drive_id(file_url: str | None) -> str | None:
    """
    Extracts the Google Drive file ID from various URL formats.

    Args:
        file_url (str | None): The URL string potentially containing a Drive file ID.

    Returns:
        str | None: The extracted file ID string, or None if no valid ID is found.
    """
    if not isinstance(file_url, str) or not file_url.strip():
        return None # Return None for empty or non-string input

    # Regex patterns to match common Google Drive URL formats containing file IDs
    # IDs are typically alphanumeric strings with underscores/hyphens, length > 20
    patterns = [
        r'/file/d/([a-zA-Z0-9_-]{20,})', # Standard /file/d/ link
        r'/d/([a-zA-Z0-9_-]{20,})',       # /d/ link (often for folders too, but ID format is key)
        r'id=([a-zA-Z0-9_-]{20,})',       # Link with id= parameter
        r'open\?id=([a-zA-Z0-9_-]{20,})'  # Older open?id= link
    ]
    for pattern in patterns:
        match = re.search(pattern, file_url)
        if match:
            file_id = match.group(1)
            # print(f"  [DEBUG] Extracted Drive ID '{file_id}' from URL.")
            return file_id

    # print(f"  [DEBUG] Could not extract Drive ID from URL: {file_url}") # Log if needed
    return None # Return None if no pattern matched


def download_drive_file(creds: Credentials, file_id: str, download_path_base: str, item_num_for_log: str = "", item_type: str = "File") -> str | None:
    """
    Downloads a file from Google Drive using its file ID.

    Handles export for Google Workspace types (Docs->DOCX, Sheets->CSV, Slides->PPTX).
    Determines the correct file extension based on metadata or export type.

    Args:
        creds (Credentials): Valid Google API credentials.
        file_id (str): The Google Drive file ID to download.
        download_path_base (str): The base path (directory + base filename without extension)
                                   for saving the downloaded file. The correct extension will be appended.
        item_num_for_log (str): Optional identifier for logging (e.g., "1", "2", "logo").
        item_type (str): Description of the item being downloaded for logs (e.g., "Image", "CSV").

    Returns:
        str | None: The full path to the successfully downloaded file (with correct extension),
                    or None if the download failed.
    """
    log_prefix = f"  {item_type}{' ' + item_num_for_log if item_num_for_log else ''}: "
    if not file_id:
        print(f"{log_prefix}[WARNING] No file ID provided for download.")
        return None

    final_download_path = None # Will store the path of the successfully saved file

    try:
        service = build('drive', 'v3', credentials=creds)
        # --- Get File Metadata ---
        # *** CORRECTED fields parameter ***
        # Request basic fields plus the 'capabilities' object which contains download/export info.
        requested_fields = 'id, name, mimeType, fileExtension, capabilities'
        file_metadata = service.files().get(
            fileId=file_id,
            fields=requested_fields
        ).execute()

        original_filename = file_metadata.get('name', f'{file_id}_download')
        mime_type = file_metadata.get('mimeType', 'application/octet-stream')
        original_extension = file_metadata.get('fileExtension', '').lower()
        # Safely access nested capabilities
        capabilities = file_metadata.get('capabilities', {})
        can_download = capabilities.get('canDownload', False)
        can_export = capabilities.get('canExport', False) # Check if export is generally possible

        print(f"{log_prefix}[INFO] File: '{original_filename}' (MIME: {mime_type}, Ext: '{original_extension}')")
        # print(f"{log_prefix}[DEBUG] Capabilities: CanDownload={can_download}, CanExport={can_export}") # Debug log

        request = None
        export_mime_type = None
        target_ext = None

        # --- Determine Download/Export Method ---
        is_google_workspace_file = 'google-apps' in mime_type
        # Check if it's a Google type AND if export is allowed by capabilities
        if is_google_workspace_file and can_export:
            # Map Google MIME types to export formats and extensions
            if 'spreadsheet' in mime_type:
                export_mime_type = 'text/csv'
                target_ext = '.csv'
                print(f"{log_prefix}[INFO] Google Sheet detected. Exporting as CSV...")
            elif 'document' in mime_type:
                 export_mime_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                 target_ext = '.docx'
                 print(f"{log_prefix}[INFO] Google Doc detected. Exporting as DOCX...")
            elif 'presentation' in mime_type:
                 export_mime_type = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
                 target_ext = '.pptx'
                 print(f"{log_prefix}[INFO] Google Slides detected. Exporting as PPTX...")
            else:
                 print(f"{log_prefix}[WARNING] Export not configured for this Google Workspace type: {mime_type}.")
                 # Fall through to check direct download capability below

            if export_mime_type:
                # Prepare export request
                request = service.files().export_media(fileId=file_id, mimeType=export_mime_type)
                final_download_path = f"{download_path_base}{target_ext}" # Define final path for export
            # else: If it's a Google type we don't explicitly export, fall through

        # --- Prepare Direct Download Request (if not exporting or if non-Google file) ---
        if request is None: # Only proceed if export wasn't prepared
            if not can_download:
                # This check is crucial for files that cannot be directly downloaded (e.g., some Google types without export)
                print(f"{log_prefix}[ERROR] Direct download is not permitted for this file (canDownload=False).")
                return None

            request = service.files().get_media(fileId=file_id)
            # Determine final path for direct downloads
            _, ext_from_name = os.path.splitext(original_filename)
            # Prioritize extension from metadata, then from original filename, else use default
            actual_ext = f".{original_extension}" if original_extension else (ext_from_name.lower() if ext_from_name else ".download")
            final_download_path = f"{download_path_base}{actual_ext}"

        # --- Execute Download/Export ---
        action_type = 'export' if export_mime_type else 'download'
        print(f"{log_prefix}[INFO] Starting {action_type} -> '{os.path.basename(final_download_path)}'...")
        fh = io.BytesIO() # Use in-memory buffer
        downloader = MediaIoBaseDownload(fh, request, chunksize=1024*1024*5) # Use chunks (e.g., 5MB)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            # Optional: print progress
            # if status: print(f"\r{log_prefix}Download {int(status.progress() * 100)}% complete.", end='')
        # print() # Newline after progress indicator finishes

        fh.seek(0) # Rewind buffer to the beginning

        # --- Save the file ---
        with open(final_download_path, 'wb') as f:
            shutil.copyfileobj(fh, f) # Efficiently write buffer to file
        print(f"{log_prefix}[SUCCESS] File saved: {os.path.basename(final_download_path)}")
        return final_download_path # Return the actual path used

    except HttpError as error:
        status = error.resp.status
        reason = error.resp.reason
        print(f'{log_prefix}[ERROR] Google Drive API HTTP error (File ID: {file_id}): Status {status} ({reason})')
        if status == 403:
            print(f'{log_prefix}  >>> Suggestion: Check file sharing permissions or API scope/enablement.')
            print(f"{log_prefix}  >>> Ensure the account running the script has access to this file.")
        elif status == 404:
            print(f'{log_prefix}  >>> Suggestion: Verify the file ID exists and is correct.')
        elif status == 400 and 'Invalid field selection' in str(error):
            # This specific error suggests the 'fields' parameter syntax needs fixing (which we did, but good to keep check)
             print(f'{log_prefix}  >>> Suggestion: Check the `fields` parameter syntax in the `files().get` API call.')
        else:
            print(f'{log_prefix}  >>> Details: {error}') # Log other HTTP errors
        traceback.print_exc() # Print stack trace for context
        return None
    except Exception as e:
        print(f'{log_prefix}[ERROR] Unexpected error during download/export for File ID {file_id}: {e}')
        traceback.print_exc()
        return None


def convert_image_if_needed(image_path: str | None, target_format: str = "PNG") -> str | None:
    """
    Converts an image to the target format (e.g., PNG) if necessary using Pillow.

    Args:
        image_path (str | None): Path to the input image file.
        target_format (str): The desired output format (e.g., "PNG", "JPEG").

    Returns:
        str | None: Path to converted image if conversion happened, else original path. None if input is None.
    """
    if image_path is None: return None

    if not PIL_INSTALLED: return image_path # Pillow needed for conversion
    if not os.path.exists(image_path):
        print(f"  [WARNING] Skipping image conversion: Input file not found: {image_path}")
        return image_path

    base_filename = os.path.basename(image_path)
    _, ext = os.path.splitext(base_filename); ext = ext.lower()
    non_image_exts = ['.csv', '.xlsx', '.xls', '.pdf', '.docx', '.txt', '.py', '.pptx', '.zip', '.json']
    if ext in non_image_exts: return image_path # Skip non-images

    target_ext = target_format.lower()
    target_path = os.path.splitext(image_path)[0] + f'.{target_ext}'

    # Avoid re-converting if target already exists and is the same format (simple check)
    # if os.path.exists(target_path) and ext == f".{target_ext}": return target_path

    try:
        with Image.open(image_path) as img:
            current_format = img.format.upper() if img.format else "UNKNOWN"
            current_mode = img.mode
            # print(f"  [DEBUG] Checking image '{base_filename}': Format={current_format}, Mode={current_mode}")

            needs_conversion = False; reason = ""
            supported_formats = ["PNG", "JPEG", "JPG", "GIF", "BMP", "TIFF", "WMF", "EMF"]
            if current_format not in supported_formats and current_format != "UNKNOWN":
                needs_conversion = True; reason = f"Format '{current_format}' not typically supported by python-docx"
            elif current_mode not in ['RGB', 'RGBA', 'L', 'P']:
                needs_conversion = True; reason = f"Mode '{current_mode}' often incompatible, converting to RGB/RGBA"
            # Additionally, force conversion if the target format is different from current detected format
            elif current_format.upper() != target_format.upper() and current_format != "UNKNOWN":
                 needs_conversion = True; reason = f"Converting from {current_format} to {target_format}"


            if needs_conversion:
                print(f"  [INFO] Converting '{base_filename}' ({reason}) to {target_format.upper()} -> '{os.path.basename(target_path)}'...")
                save_kwargs = {}; img_to_save = img

                # Handle mode conversions before saving
                if img_to_save.mode == 'P':
                    if 'transparency' in img_to_save.info: img_to_save = img_to_save.convert('RGBA')
                    else: img_to_save = img_to_save.convert('RGB')
                elif img_to_save.mode == 'LA': img_to_save = img_to_save.convert('RGBA')
                elif img_to_save.mode == 'CMYK': img_to_save = img_to_save.convert('RGB')
                elif img_to_save.mode not in ['RGB', 'RGBA', 'L']: img_to_save = img_to_save.convert('RGB')

                # Handle target format specifics (e.g., JPEG transparency)
                if target_format.upper() == 'JPEG':
                    save_kwargs['quality'] = 95
                    if img_to_save.mode == 'RGBA':
                       print("    [INFO] Converting RGBA to RGB for JPEG (pasting on white background)...")
                       bg = Image.new('RGB', img_to_save.size, (255, 255, 255))
                       try: bg.paste(img_to_save, mask=img_to_save.split()[3]) # Use alpha mask
                       except: bg.paste(img_to_save, (0,0)) # Fallback if alpha split fails
                       img_to_save = bg

                img_to_save.save(target_path, format=target_format.upper(), **save_kwargs)
                print(f"  [SUCCESS] Converted image saved: '{os.path.basename(target_path)}'")
                return target_path
            else:
                return image_path # Return original path if no conversion needed

    except UnidentifiedImageError: print(f"  [WARNING] Pillow could not identify format for '{base_filename}'. Cannot convert."); return image_path
    except FileNotFoundError: print(f"  [ERROR] Image file not found during conversion: {image_path}"); return image_path
    except Exception as e: print(f"  [ERROR] Failed to convert image '{base_filename}': {e}"); traceback.print_exc(); return image_path


def sanitize_filename(name: str, max_length: int = 100) -> str:
    """ Cleans and sanitizes a string for use as part of a filename. """
    if not isinstance(name, str): name = str(name)
    sanitized = re.sub(r'[\\/*?:"<>|]', "", name) # Remove invalid chars
    sanitized = re.sub(r'\s+|\.+', '_', sanitized) # Replace whitespace/dots with underscore
    sanitized = re.sub(r'_+', '_', sanitized) # Collapse multiple underscores
    sanitized = sanitized.strip('_ ') # Trim leading/trailing underscores/spaces
    sanitized = sanitized[:max_length] # Truncate
    if not sanitized: return "Untitled_Report" # Fallback name
    return sanitized


# ===========================================================
# --- Main Processing Logic ---
# ===========================================================

def main():
    """ Main function to orchestrate the report generation and emailing process. """
    print(f"=== {get_timestamp()} Report Generation Process Starting ===")
    main_start_time = time.time(); run_status = "SUCCESS"

    # --- Setup ---
    print("\n[INFO] Performing initial setup...")
    try: # Ensure essential folders exist
        for folder in [config.OUTPUT_FOLDER, config.TEMP_FOLDER, config.TEMPLATE_FOLDER]:
            if not os.path.exists(folder): print(f"  [INFO] Creating missing folder: {folder}"); os.makedirs(folder)
        print("[SUCCESS] Folder structure verified/created.")
    except OSError as e: print(f"[ERROR] Failed to create required folders: {e}"); sys.exit("Exiting.")

    # --- Authenticate ---
    creds = authenticate()
    if not creds: sys.exit("Exiting: Google Authentication failed.")

    # --- Get Sheet Data ---
    headers, responses = get_sheet_data(creds)
    if headers is None or responses is None: sys.exit("Exiting: Failed to retrieve sheet data.")
    if not responses: print("[INFO] No data rows found."); print(f"\n=== {get_timestamp()} Finished (No data) ==="); return

    # --- Validate Core Headers ---
    print("\n[INFO] Validating essential sheet headers...")
    core_headers = [config.HEADER_TIMESTAMP, config.HEADER_NAME, config.HEADER_EMAIL_RECIPIENT, config.HEADER_TEMPLATE_CHOICE]
    missing_core_headers = [h for h in core_headers if h not in headers]
    if missing_core_headers: print(f"[ERROR] Core header(s) missing: {missing_core_headers}"); sys.exit("Exiting.")
    else: print("[SUCCESS] Core headers validated.")

    # --- Map Template Choices to Report Modules ---
    report_modules_map = { # lowercase keys
        'event report': event_report, 'remedial class report': remedial_report,
        'remedial class': remedial_report, 'curriculum map': curriculum_map_report,
        'summarise report': report_summarizer, 'summarize report': report_summarizer,
    }
    print(f"[INFO] Registered report types: {list(report_modules_map.keys())}")

    # --- Process Each Response Row ---
    processed_count, emailed_count, skipped_count, error_count = 0, 0, 0, 0
    temp_files_overall = [] # Tracks temp files from successfully emailed reports
    total_rows = len(responses)
    print(f"\n=== {get_timestamp()} Processing {total_rows} Response Rows ===")

    for i, row_data_raw in enumerate(responses):
        row_start_time = time.time()
        row_num_for_log = i + 2 # Sheet row number
        print(f"\n--- {get_timestamp()} Processing Sheet Row {row_num_for_log}/{total_rows+1} ---")
        current_row_temp_files, current_report_assets = [], {}
        generated_doc, report_module = None, None
        save_successful, email_successful = False, False
        row_status = "PENDING"

        try:
            # --- Extract Row Data & Key Info ---
            if len(row_data_raw) < len(headers): row_data_raw.extend([''] * (len(headers) - len(row_data_raw)))
            response_dict = {headers[j]: row_data_raw[j] for j in range(len(headers))}
            timestamp_str = response_dict.get(config.HEADER_TIMESTAMP, f'Row_{row_num_for_log}')
            recipient_email = response_dict.get(config.HEADER_EMAIL_RECIPIENT, '').strip()
            template_choice_raw = response_dict.get(config.HEADER_TEMPLATE_CHOICE, '').strip()
            submitter_name = response_dict.get(config.HEADER_NAME, 'UnknownSubmitter').strip()
            template_choice_processed = template_choice_raw.lower().strip()

            print(f"  [INFO] Submitter: {submitter_name}")
            print(f"  [INFO] Recipient: {recipient_email if recipient_email else '[Not Provided]'}")
            print(f"  [INFO] Template Choice: '{template_choice_raw}'")

            if not template_choice_processed: print("  [WARNING] Skipping row - Template choice missing."); skipped_count += 1; row_status = "SKIPPED"; continue
            report_module = report_modules_map.get(template_choice_processed)
            if not report_module: print(f"  [WARNING] Skipping row - No implementation for template: '{template_choice_raw}'"); skipped_count += 1; row_status = "SKIPPED"; continue

            # --- Prepare Filename ---
            report_title_base, date_str = 'Report', ''
            try: # Robust date parsing
                dt_object = None; possible_formats = ['%m/%d/%Y %H:%M:%S', '%m/%d/%Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y %H:%M:%S', '%d-%b-%Y %H:%M:%S']
                for fmt in possible_formats:
                    try: dt_object = datetime.strptime(timestamp_str, fmt); break
                    except ValueError: continue
                if dt_object: date_str = dt_object.strftime('%Y%m%d')
                else: raise ValueError("Date format not recognized")
            except ValueError: date_str = re.sub(r'[\W]+', "_", timestamp_str.split(' ')[0])[:10]; print(f"    [WARNING] Using fallback date string: '{date_str}'")
            safe_submitter_name = sanitize_filename(submitter_name, 30)

            if template_choice_processed == 'event report': report_title_base = sanitize_filename(response_dict.get(config.HEADER_EVENT_TITLE, 'Event'), 50)
            elif template_choice_processed in ['remedial class report', 'remedial class']: report_title_base = f"Remedial_{sanitize_filename(response_dict.get(config.HEADER_REMEDIAL_COURSE, 'Course'), 40)}"
            elif template_choice_processed == 'curriculum map': report_title_base = f"CurriculumMap_{sanitize_filename(response_dict.get(config.HEADER_CM_EVENT_NAME, 'Event'), 40)}"
            elif template_choice_processed in ['summarise report', 'summarize report']: report_title_base = f"Summary"
            final_report_filename_base = f"{report_title_base}_{date_str}_{safe_submitter_name}"
            print(f"  [INFO] Base report filename: {final_report_filename_base}")

            # --- Download Assets & Generate Report ---
            print(f"\n  [INFO] Starting assets/generation for '{template_choice_raw}'...")
            if report_module == event_report: # --- Event Report ---
                 if not os.path.exists(report_module.EVENT_REPORT_TEMPLATE_PATH): raise FileNotFoundError(f"Template missing: {report_module.EVENT_REPORT_TEMPLATE_PATH}")
                 for idx, header in enumerate([config.HEADER_EVENT_IMAGE_1, config.HEADER_EVENT_IMAGE_2]):
                     url = response_dict.get(header); file_id = extract_drive_id(url)
                     if file_id:
                          base = os.path.join(config.TEMP_FOLDER, f"event_img{idx+1}_{final_report_filename_base}_{file_id}")
                          dl_path = download_drive_file(creds, file_id, base, str(idx+1), "Image")
                          if dl_path: converted_path = convert_image_if_needed(dl_path, "PNG"); current_report_assets[f'image{idx+1}'] = converted_path; current_row_temp_files.append(dl_path); (lambda p: current_row_temp_files.append(p) if p != dl_path else None)(converted_path)
                 generated_doc = report_module.generate_report(response_dict, current_report_assets.get('image1'), current_report_assets.get('image2'))

            elif report_module == remedial_report: # --- Remedial Report ---
                 if not os.path.exists(report_module.REMEDIAL_REPORT_TEMPLATE_PATH): raise FileNotFoundError(f"Template missing: {report_module.REMEDIAL_REPORT_TEMPLATE_PATH}")
                 csv_headers = {'csv1': config.HEADER_REMEDIAL_STUDENT_LIST_CSV, 'csv2': config.HEADER_REMEDIAL_TIMETABLE_CSV, 'csv3': config.HEADER_REMEDIAL_ATTENDANCE_CSV}
                 for key, header in csv_headers.items():
                      url = response_dict.get(header); file_id = extract_drive_id(url)
                      if file_id: base = os.path.join(config.TEMP_FOLDER, f"remedial_{key}_{final_report_filename_base}_{file_id}"); dl_path = download_drive_file(creds, file_id, base, key, "CSV"); current_report_assets[key] = dl_path; current_row_temp_files.append(dl_path) if dl_path else None
                 generated_doc = report_module.generate_report(response_dict, current_report_assets.get('csv1'), current_report_assets.get('csv2'), current_report_assets.get('csv3'))

            elif report_module == curriculum_map_report: # --- Curriculum Map ---
                 if not os.path.exists(report_module.CURRICULUM_MAP_REPORT_TEMPLATE_PATH): raise FileNotFoundError(f"Template missing: {report_module.CURRICULUM_MAP_REPORT_TEMPLATE_PATH}")
                 assets_cfg = {'logo': (config.HEADER_CM_CLUB_LOGO, "Image", "PNG"), 'guests_csv': (config.HEADER_CM_GUESTS_CSV, "CSV", None),'program_csv': (config.HEADER_CM_PROGRAM_CSV, "CSV", None), 'lecture_plan_csv': (config.HEADER_CM_LECTURE_PLAN_CSV, "CSV", None),'photos': (config.HEADER_CM_PHOTOS, "Image", "PNG"), 'brochure': (config.HEADER_CM_BROCHURE, "Image", "PNG"),'schedule_csv': (config.HEADER_CM_SCHEDULE_CSV, "CSV", None), 'attendance_csv': (config.HEADER_CM_ATTENDANCE_CSV, "CSV", None),'news': (config.HEADER_CM_NEWS, "File", None)}
                 for key, (header, asset_type, img_fmt) in assets_cfg.items():
                     url = response_dict.get(header); file_id = extract_drive_id(url)
                     if file_id: base = os.path.join(config.TEMP_FOLDER, f"cm_{key}_{final_report_filename_base}_{file_id}"); dl_path = download_drive_file(creds, file_id, base, key, asset_type);
                     if dl_path: final_path = convert_image_if_needed(dl_path, img_fmt) if asset_type=="Image" and img_fmt else dl_path; current_report_assets[key] = final_path; current_row_temp_files.append(dl_path); (lambda p: current_row_temp_files.append(p) if p != dl_path else None)(final_path)
                 generated_doc = report_module.generate_report(response_dict, current_report_assets)

            elif report_module == report_summarizer: # --- Summarizer Report ---
                 report_url = response_dict.get(config.HEADER_SUMMARY_UPLOAD_FILE); file_id = extract_drive_id(report_url)
                 if not file_id: raise ValueError("Summarize Report requires an uploaded file link.")
                 base = os.path.join(config.TEMP_FOLDER, f"summary_input_{final_report_filename_base}_{file_id}")
                 local_report_path = download_drive_file(creds, file_id, base, "", "ReportToSummarize")
                 if not local_report_path or not os.path.exists(local_report_path): raise FileNotFoundError(f"Failed to download input report file (ID: {file_id}) for summarization.")
                 current_row_temp_files.append(local_report_path); current_report_assets['report_to_summarize'] = local_report_path
                 generated_doc, extracted_image_paths = report_module.generate_report(response_dict, current_report_assets)
                 if extracted_image_paths: current_row_temp_files.extend(extracted_image_paths)

            # --- Check Document Generation Success ---
            if generated_doc is None: raise RuntimeError(f"Generation failed for '{template_choice_raw}' (module returned None).")
            print(f"  [SUCCESS] Content generation complete for '{template_choice_raw}'.")

            # --- Save Report ---
            report_filename = f"{final_report_filename_base}.docx"
            report_filepath = os.path.join(config.OUTPUT_FOLDER, report_filename)
            print(f"  [INFO] Saving generated report to: {report_filepath}")
            generated_doc.save(report_filepath); save_successful = True
            print(f"  [SUCCESS] Report saved: {os.path.basename(report_filepath)}"); processed_count += 1

            # --- Email Report ---
            if recipient_email:
                print(f"\n  [INFO] Attempting to email report to {recipient_email}...")
                email_successful = email_sender.send_email_with_attachment(recipient_email, submitter_name, template_choice_raw, timestamp_str, report_filepath)
                if email_successful: emailed_count += 1; temp_files_overall.extend(current_row_temp_files); # print(f"    [DEBUG] Marked {len(current_row_temp_files)} temp file(s) for cleanup.")
                else: print(f"    [WARNING] Email sending failed. Temp files for row {row_num_for_log} will NOT be cleaned up."); row_status = "ERROR (Email Failed)"
            else: print("\n  [INFO] No recipient email provided. Skipping email. Temp files will NOT be cleaned up."); row_status = "SUCCESS (Not Emailed)"


        # --- Catch Errors During Row Processing ---
        except (FileNotFoundError, ValueError, RuntimeError, HttpError, ImportError) as row_err:
             err_type = type(row_err).__name__
             print(f"  [ERROR] {err_type}: {row_err}")
             if isinstance(row_err, HttpError): traceback.print_exc(limit=1) # Short traceback for API errors
             elif not isinstance(row_err, FileNotFoundError) and not isinstance(row_err, ValueError): traceback.print_exc() # Full traceback for others
             row_status = f"ERROR ({err_type})"
             error_count += 1
        except Exception as row_err: # Catch any other unexpected error
            print(f"  [ERROR] An unexpected error occurred processing row {row_num_for_log}: {row_err}")
            traceback.print_exc()
            row_status = "ERROR (Unexpected)"; error_count += 1

        # --- Row Post-Processing & Immediate Cleanup on Failure/Skip ---
        if row_status == "PENDING": row_status = "SUCCESS" # If no errors caught, mark as success
        row_end_time = time.time()
        print(f"--- Row {row_num_for_log} Processing Finished ({row_end_time - row_start_time:.2f}s). Status: [{row_status}] ---")

        if not row_status.startswith("SUCCESS") or (row_status == "SUCCESS" and not email_successful and recipient_email): # Clean up if error OR if not successfully emailed (when email was expected)
            if current_row_temp_files:
                 cleanup_reason = "error occurred" if row_status.startswith("ERROR") else ("skipped" if row_status == "SKIPPED" else "not emailed successfully")
                 print(f"  [INFO] Cleaning up {len(current_row_temp_files)} temp file(s) for row {row_num_for_log} (Reason: {cleanup_reason}).")
                 cleaned_now = 0
                 for f_path in set(f for f in current_row_temp_files if f and os.path.exists(f)):
                     try: os.remove(f_path); cleaned_now += 1
                     except OSError as e_clean: print(f"    [WARNING] Could not delete temp file '{f_path}': {e_clean}")
                 # print(f"    [DEBUG] Cleaned up {cleaned_now} files for this row.")

        # End of row processing loop

    # ===========================================================
    # --- Final Summary and Cleanup ---
    # ===========================================================
    main_end_time = time.time()
    total_failed_or_skipped = total_rows - processed_count
    if error_count > 0 or skipped_count > 0: run_status = "COMPLETED WITH ISSUES"

    print(f"\n\n=== {get_timestamp()} Report Generation Process Finished ===")
    print(f"  Overall Status: [{run_status}]")
    print(f"  Total Run Time: {main_end_time - main_start_time:.2f} seconds")
    print("-" * 50)
    print("  PROCESSING SUMMARY:")
    print(f"  - Total Responses Found:         {total_rows}")
    print(f"  - Reports Generated & Saved:     {processed_count}")
    print(f"  - Reports Successfully Emailed:    {emailed_count}")
    print(f"  - Rows Skipped or Failed:      {total_failed_or_skipped}")
    if error_count > 0:   print(f"    - Errors Encountered:          {error_count}")
    if skipped_count > 0: print(f"    - Intentionally Skipped Rows:  {skipped_count}")
    print("-" * 50)

    # --- Final Cleanup of Temp Files (from successfully emailed reports) ---
    if temp_files_overall:
        unique_temp_files_to_clean = set(f for f in temp_files_overall if f) # Use set, filter None
        print(f"\n[INFO] Starting final cleanup of {len(unique_temp_files_to_clean)} temporary file(s) from successfully emailed reports...")
        cleaned_count, failed_cleanup_count = 0, 0
        for f_path in unique_temp_files_to_clean:
             if os.path.exists(f_path):
                 try: os.remove(f_path); cleaned_count += 1
                 except OSError as e: print(f"  [WARNING] Could not delete temp file '{f_path}': {e}"); failed_cleanup_count += 1
        print(f"[INFO] Final cleanup complete. Deleted {cleaned_count} file(s).")
        if failed_cleanup_count > 0: print(f"  [WARNING] Failed to delete {failed_cleanup_count} temp file(s). Manual check of '{config.TEMP_FOLDER}' recommended.")
    else:
        print("\n[INFO] No temporary files marked for final cleanup.")

    # Optional: Attempt to remove the temp folder *only if it's empty*
    try:
        if os.path.exists(config.TEMP_FOLDER) and not os.listdir(config.TEMP_FOLDER):
            os.rmdir(config.TEMP_FOLDER); print(f"[INFO] Removed empty temporary folder: {config.TEMP_FOLDER}")
    except OSError: pass # Ignore errors (e.g., dir not empty, permissions)

    print(f"\n=== {get_timestamp()} Script Finished. ===")

# --- Script Entry Point ---
if __name__ == '__main__':
    if not GOOGLE_LIBS_AVAILABLE or not MODULES_AVAILABLE: sys.exit("Exiting: Prerequisite libraries or modules missing.")
    main()