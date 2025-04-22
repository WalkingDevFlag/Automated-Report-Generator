# Automated Report Generator (R_K_Tailor Project)

## Overview

## Setup and Installation

1.  **Clone the Repository:**
    ```bash
    git clone https://github.com/WalkingDevFlag/R_K_Tailor.git
    cd R_K_Tailor
    ```
2.  **Create a Python Environment (Recommended):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows use `venv\Scripts\activate`
    ```
3.  **Install Dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
4.  **Google Cloud Platform Setup:**
    *   Create a new project on the [Google Cloud Console](https://console.cloud.google.com/).
    *   **Enable APIs:** Enable the "Google Drive API" and "Google Sheets API" for your project.
    *   **Create OAuth 2.0 Credentials:**
        *   Go to "APIs & Services" -> "Credentials".
        *   Click "Create Credentials" -> "OAuth client ID".
        *   Select "Desktop app" as the Application type.
        *   Give it a name (e.g., "Report Generator Script").
        *   Click "Create".
    *   **Download Credentials:** Download the JSON file containing your client ID and secret. Rename this file to `credentials.json` and place it in the project's root directory. **Do not commit this file to version control.**
5.  **Create `.env` File:**
    *   Create a file named `.env` in the project's root directory.
    *   Add the following variables, replacing the placeholder values:
        ```dotenv
        # .env file
        SPREADSHEET_ID=YOUR_GOOGLE_SHEET_ID_HERE
        EMAIL_SENDER=your_sender_email@gmail.com
        EMAIL_PASSWORD=your_gmail_app_password_here
        ```
    *   `SPREADSHEET_ID`: The ID of the Google Sheet collecting the form responses (from the sheet URL).
    *   `EMAIL_SENDER`: The Gmail address the reports will be sent *from*.
    *   `EMAIL_PASSWORD`: **Important:** If using Gmail and you have 2-Factor Authentication enabled, you **must** generate an "App Password" for this script. Do not use your regular Gmail password here. If 2FA is off, your regular password might work, but App Passwords are more secure.

## Configuration

Besides the `.env` file, review `config.py` for other settings:

*   **Google Sheet:**
    *   `SHEET_NAME`: Ensure this matches the name of the sheet tab containing responses.
    *   `LAST_COLUMN`: Adjust if your sheet uses more columns than `AZ`.
    *   `RANGE_NAME`: Automatically constructed from `SHEET_NAME` and `LAST_COLUMN`.
    *   `HEADER_...`: **Crucially important!** These constants *must exactly match* the column headers in the *first row* of your Google Sheet. Update them in `config.py` if your form/sheet headers are different.
*   **File Paths:** `OUTPUT_FOLDER`, `TEMP_FOLDER`, `TEMPLATE_FOLDER` can be adjusted if needed.
*   **Email:** `SMTP_SERVER`, `SMTP_PORT` might need changing if not using Gmail.
*   **Report Modules:** Specific configurations (template filenames, placeholder names, image sizes, summarizer model) are within each report module file (e.g., `curriculum_map_report.py`).
*   **Templates:** Ensure your `.docx` files in the `template/` directory contain the correct placeholder text (`{{PlaceholderName}}`, `{{table1}}`, etc.) that corresponds to the mappings in `config.py` and the logic within the report modules. The provided `curriculum_map_report.docx` serves as the template for that specific report type.

## Usage

1.  Ensure you have completed the Setup and Configuration steps.
2.  Activate your Python environment (if used).
3.  Run the main script from the project's root directory:
    ```bash
    python main.py
    ```
4.  **First Run:** The script will likely open a web browser window asking you to log in to your Google account and grant the script permission to access your Google Sheets and Drive data. Follow the prompts to authorize access. After successful authorization, a `token.json` file will be created, and the script will proceed. Subsequent runs should use the `token.json` file automatically unless it expires or is deleted.
5.  The script will then process each row in the Google Sheet, generating and emailing reports as configured. Check the console output for progress and any errors. Generated reports will appear in the `generated_reports` folder.

## Supported Report Types

*   **Event Report (`event_report.py`):** Simple report using text and image placeholders.
*   **Remedial Class Report (`remedial_report.py`):** Inserts data into text placeholders and creates tables from linked CSV files (`{{table1}}`, `{{table2}}`, `{{table3}}`). Uses run-level text replacement for better formatting preservation.
*   **Curriculum Map Report (`curriculum_map_report.py`):** More complex report with dynamic section inclusion/exclusion based on provided data (text, images, CSVs). It automatically renumbers sections and the table of contents, inserts tables and images, and enforces font styles. Requires a specific template structure (like the provided `curriculum_map_report.docx`).
*   **Summarizer Report (`report_summarizer.py`):** Takes a linked PDF or DOCX file, extracts text and images, uses a Hugging Face transformer model to summarize the text, and creates a *new* DOCX file containing the summary and images. Requires `PyMuPDF`, `transformers`, `torch`, `accelerate`, and `Pillow`.

## Dependencies

All required Python libraries are listed in `requirements.txt`. Key dependencies include:

*   `google-api-python-client`, `google-auth-httplib2`, `google-auth-oauthlib`: For Google API interaction.
*   `python-docx`: For reading and writing DOCX files.
*   `docxtpl`: Used by Curriculum Map report for some template operations (like image replacement).
*   `python-dotenv`: For loading environment variables.
*   `Pillow`: For image processing and conversion (Optional but recommended).
*   `PyMuPDF`: For extracting text/images from PDF files (Required for Summarizer).
*   `transformers`, `torch`, `huggingface_hub`, `accelerate`: For the text summarization feature (Required for Summarizer).

## Troubleshooting

*   **Authentication Errors / `token.json` issues:**
    *   **Problem:** Script fails during authentication, browser doesn't open, or `token.json` seems corrupted.
    *   **Solution:** Delete the `token.json` file and re-run `main.py`. This will force the OAuth 2.0 flow again. Ensure `credentials.json` is present and correct. Verify the redirect URIs in your Google Cloud Console credential settings match what the `google-auth-oauthlib` expects (usually `http://localhost:<port>`).
*   **`ModuleNotFoundError`:**
    *   **Problem:** Script fails immediately, reporting a missing module.
    *   **Solution:** Ensure you have activated your virtual environment (if used) and successfully installed all packages using `pip install -r requirements.txt`.
*   **Sheet Data Not Found / Incorrect Headers:**
    *   **Problem:** Script reports "No data found" or fails with errors related to missing keys/headers.
    *   **Solution:** Double-check `SPREADSHEET_ID`, `SHEET_NAME`, and `RANGE_NAME` in `config.py`. **Crucially**, verify that the `HEADER_...` constants in `config.py` *exactly* match the column headers in the first row of your Google Sheet (case-sensitive, including spaces).
*   **File Download Errors (403 Forbidden):**
    *   **Problem:** Script fails to download a file from Google Drive, reporting a 403 error.
    *   **Solution:** The account used to authenticate the script does not have permission to access the file. Ensure the file(s) linked in the Google Form response are shared appropriately (e.g., "Anyone with the link can view" or shared directly with the email account used for authentication).
*   **File Download Errors (404 Not Found):**
    *   **Problem:** Script fails to download a file, reporting a 404 error.
    *   **Solution:** The link in the Google Form response is likely incorrect, or the file has been deleted. Verify the link and the existence of the file on Google Drive. Check if `extract_drive_id` is correctly parsing the ID from the URL format used.
*   **Image/Table Insertion Issues:**
    *   **Problem:** Images or tables are missing, appear in the wrong place, or placeholder text (`{{table1}}`, `{{EventBrochure}}`) remains.
    *   **Solution:**
        *   Verify the placeholder tags in your `.docx` template *exactly* match the tags defined in the relevant report module config (e.g., `event_report.py`, `remedial_report.py`).
        *   For table insertion in Remedial/Curriculum Map reports, ensure the placeholder tag is in its *own separate paragraph* immediately following any related heading.
        *   Ensure the CSV/Image files were downloaded successfully to `temp_files`. Check console logs for download errors.
        *   For images, ensure Pillow is installed (`pip install Pillow`) if using formats that need conversion (like WEBP).
*   **Email Sending Fails:**
    *   **Problem:** Reports generate but are not emailed. Errors related to SMTP authentication or connection appear.
    *   **Solution:**
        *   Verify `EMAIL_SENDER` and `EMAIL_PASSWORD` in your `.env` file are correct.
        *   If using Gmail with 2FA, ensure you are using a generated **App Password**, not your regular account password.
        *   Check Gmail security settings – less secure app access might need to be enabled (though App Passwords are the preferred method).
        *   Verify `SMTP_SERVER` and `SMTP_PORT` in `config.py` are correct for your provider (Gmail defaults are usually `smtp.gmail.com` and `465` for SSL or `587` for TLS). Ensure `config.EMAIL_USE_SSL` matches the port choice.
        *   Check network/firewall settings that might block SMTP connections.
*   **Summarizer Errors (CUDA Out of Memory, etc.):**
    *   **Problem:** Summarizer report fails during the `summarize_text` step.
    *   **Solution:** See suggestions in `report_summarizer.py`'s `summarize_text` function docstring/comments. Common fixes include: trying a smaller model (edit `SUMMARIZER_MODEL`), reducing `MANUAL_TRUNCATE_TOKEN_LIMIT`, forcing CPU use (`FORCE_CPU = True`), or ensuring sufficient system RAM/VRAM and compatible ML library versions (`transformers`, `torch`, `accelerate`).

## Security Considerations

*   **`credentials.json`:** This file contains your OAuth 2.0 client secrets. **Keep it confidential.** Do not commit it to version control (add it to your `.gitignore` file).
*   **`.env` File:** This file contains your spreadsheet ID and email credentials. **Keep it confidential.** Do not commit it to version control (add it to your `.gitignore` file).
*   **`token.json`:** This file contains access and refresh tokens granting the script access to your Google account (within the requested scopes). While less sensitive than your primary password, it should also be kept secure and **not committed to version control** (add it to `.gitignore`).
*   **Email Credentials:** Use an App Password if using Gmail with 2FA. Avoid using your main account password directly in the `.env` file if possible.
*   **Scopes:** The script requests `readonly` access for Sheets and Drive (`config.SCOPES`). Avoid requesting broader permissions unless strictly necessary.

## Known Issues

*   **Empty Form Fields:** As observed, if a field in the Google Form intended for a placeholder (e.g., `{{SomeOptionalField}}`) is left empty in a submission, the generation script might not reliably remove the placeholder tag itself from the final document. While the code generally attempts to replace placeholders with the provided value (which would be an empty string for an empty field), some complex template structures or specific placeholder handling logic (especially in the Curriculum Map report's dynamic section management *before* cleanup) might result in the literal placeholder text remaining visible. It's recommended to review generated reports, especially when optional fields are omitted.
