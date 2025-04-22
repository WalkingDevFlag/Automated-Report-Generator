# Automated Report Generator from Google Forms

[![Python Version](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT) <!-- Optional License Badge -->

This project provides a robust solution for automatically generating formatted Microsoft Word (.docx) reports based on user submissions via a Google Form. It retrieves data from a linked Google Sheet, populates predefined templates, handles image uploads, converts image formats, and emails the final report.

## Features

*   **Google Form/Sheet Integration:** Processes responses collected in a Google Sheet.
*   **Multi-Template Support:** Handles different report types selected by the user in the form (easily extensible).
*   **Dynamic Content Filling:** Populates placeholders (`{{placeholder}}`) in `.docx` templates with corresponding sheet data.
*   **Image Handling:** Downloads images uploaded via the form from Google Drive using API calls.
*   **Image Conversion:** Includes capability (using Pillow) to convert common unsupported formats (like WEBP) to PNG for reliable insertion into `.docx` files.
*   **Automated Emailing:** Sends the generated report as an attachment to a specified recipient email address using SMTP.
*   **Modular & Configurable:**
    *   Separates concerns into distinct Python modules (`main`, `config`, `email_sender`, report-specific modules like `event_report`).
    *   Uses a `.env` file for secure storage of sensitive credentials (API keys, passwords).
    *   Centralizes configuration (sheet ID, range, headers, paths) in `config.py`.
*   **Descriptive Filenaming:** Generates meaningful report filenames using key data like report title, date, and submitter name.
*   **Conditional Cleanup:** Automatically removes temporary image files only after the corresponding report has been successfully generated and emailed.

## Project Structure

```
.
├── .env                     # Store sensitive credentials here (Create this file)
├── .gitignore               # Specifies intentionally untracked files
├── config.py                # Loads .env, defines configurations & sheet headers
├── email_sender.py          # Module for sending emails via SMTP
├── event_report.py          # Module for generating 'Event Report' type
├── main.py                  # Main script orchestrating the workflow
├── requirements.txt         # Python package dependencies
├── template/
│   └── event_report.docx    # Example Word template (Add others here)
├── credentials.json         # Google OAuth 2.0 Client ID file (From GCP)
├── token.json               # Google API token (Generated on first run)
├── generated_reports/       # Default output folder for generated reports
└── temp_images/             # Default folder for temporary image downloads
```

## Setup and Installation

1.  **Prerequisites:**
    *   Python 3.7+
    *   Google Account
    *   Google Cloud Platform (GCP) Project

2.  **Clone Repository:**
    ```bash
    git clone https://github.com/WalkingDevFlag/Automated-Report-Generator.git
    cd Automated-Report-Generator
    ```

3.  **Create Virtual Environment (Recommended):**
    ```bash
    python -m venv venv
    # Activate:
    # Windows: .\venv\Scripts\activate
    # macOS/Linux: source venv/bin/activate
    ```

4.  **Install Dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

5.  **Google Cloud Platform Setup:**
    *   Go to your GCP Console.
    *   Create a new project or select an existing one.
    *   **Enable APIs:** Enable "Google Sheets API" and "Google Drive API".
    *   **Create OAuth Credentials:**
        *   Navigate to "APIs & Services" > "Credentials".
        *   Click "+ CREATE CREDENTIALS" > "OAuth client ID".
        *   Select Application type: "Desktop app".
        *   Give it a name (e.g., "Report Generator Script").
        *   Click "Create".
        *   **Download the JSON file.** Rename it to `credentials.json` and place it in the project's root directory.
    *   **Configure OAuth Consent Screen:**
        *   Navigate to "APIs & Services" > "OAuth consent screen".
        *   Choose "External" user type (unless restricted to a Google Workspace).
        *   Fill in required fields (App name, User support email, Developer contact).
        *   **IMPORTANT:** Keep the "Publishing status" as **"Testing"** initially. While in testing mode, you **must add the Google account(s) that will run the script** under the "Test users" section. Failure to do so will result in authentication errors ("Access blocked").

6.  **Google Form & Sheet Setup:**
    *   Create your Google Form. Include questions for:
        *   Submitter's Name
        *   Recipient Email Address
        *   Report Template Selection (Dropdown/Multiple Choice)
        *   Template-specific data fields
        *   File Upload fields for necessary images.
    *   Link the form responses to a new Google Sheet.
    *   **Verify Headers:** Ensure the column headers in the first row of the response sheet **exactly match** the `HEADER_` constants defined in `config.py` (case and spacing matter!).
    *   Note the **Spreadsheet ID** from the Sheet URL.
    *   Note the **Sheet Tab Name** (e.g., `Form Responses 1`).

7.  **Word Template(s) (`template/` folder):**
    *   Create/edit your `.docx` template files (e.g., `event_report.docx`).
    *   Use placeholders like `{{PlaceholderName}}` for dynamic text. These **must exactly match** the keys in the `_PLACEHOLDERS` dictionary in the corresponding report module (e.g., `event_report.py`).
    *   Use specific tags like `{{EventBrochure}}` (verify exact text including spaces!) where images should be inserted. These **must exactly match** the `IMAGE_PLACEHOLDER_TAG_` constants in the report module.
    *   Place templates in the `template/` directory.

8.  **Create and Configure `.env` File:**
    *   Create a file named `.env` in the project root.
    *   Add and configure the following variables:
        ```dotenv
        SPREADSHEET_ID="YOUR_GOOGLE_SHEET_ID" # Get from Sheet URL
        EMAIL_SENDER="your_sending_email@example.com"
        EMAIL_PASSWORD="YOUR_APP_PASSWORD_OR_EMAIL_PASSWORD" # Use Gmail App Password!
        ```
    *   Add `.env` to your `.gitignore` file if it's not already there.

9.  **Configure `config.py`:**
    *   Double-check `RANGE_NAME` to match your response sheet tab name and column range (e.g., `Form Responses 1!A:M`).
    *   Verify all `HEADER_` constants match your sheet headers exactly.
    *   Adjust `SMTP_SERVER` / `SMTP_PORT` if not using Gmail default SSL.

10. **Configure Report Module(s) (e.g., `event_report.py`):**
    *   Verify `_PLACEHOLDERS` dictionary keys match template placeholders.
    *   Verify `IMAGE_PLACEHOLDER_TAG_` constants match template image tags *exactly*.

## How to Run

1.  **Activate Virtual Environment** (if applicable).
2.  **Navigate** to the project directory in your terminal.
3.  **Run the main script:**
    ```bash
    python main.py
    ```
4.  **First Run Only:** A browser window will open for Google Authentication. Log in with an authorized Google account (one listed as a Test User if applicable) and grant permissions. A `token.json` file will be created for subsequent runs.
5.  **Monitor Output:** Check the console for progress, generated report paths, email status, and any errors. Reports are saved in `generated_reports/`.

## Adding New Report Templates

1.  **Create `.docx` Template:** Design the new template (e.g., `curriculum_map.docx`) with placeholders/tags and save it in `template/`.
2.  **Create Python Module:** Create `curriculum_map.py`.
    *   Define `TEMPLATE_FILENAME`, `_PLACEHOLDERS` mapping (using headers from `config.py`), and `IMAGE_PLACEHOLDER_TAG_` constants specific to this template.
    *   Implement a `generate_report(data, image1_local_path=None, image2_local_path=None)` function.
3.  **Update `config.py`:**
    *   Add any new `HEADER_` constants for columns unique to this template.
    *   Ensure `RANGE_NAME` covers these columns.
4.  **Update `main.py`:**
    *   `import curriculum_map`
    *   Add an `elif processed_template_choice == 'curriculum map':` block in the "Select Report Module" section to set `report_module = curriculum_map` and assign relevant `image_header_` variables from `config.py`.

## Troubleshooting Common Issues

*   **Authentication Errors (`access_denied`, `invalid_grant`):** Ensure you are using an authorized Test User (if app is "Testing"). Delete `token.json` and re-run `main.py` to re-authenticate. Check `credentials.json`.
*   **`SMTP Authentication failed`:** Wrong `EMAIL_PASSWORD` in `.env`. **Use a Gmail App Password** for 2FA accounts. Check `EMAIL_SENDER`. Check Google security alerts for blocked sign-ins.
*   **`Column header '...' not found`:** Header in `config.py` doesn't match the sheet header exactly. Check spelling, case, and spaces.
*   **Placeholders Not Replaced:** Mismatch between `{{placeholder}}` in `.docx` and keys in the report module's `_PLACEHOLDERS` dictionary. Check exact text.
*   **Images Not Inserted / Errors:**
    *   Mismatch between image tags (`{{...}}`) in `.docx` and `IMAGE_PLACEHOLDER_TAG_` constants in the report module. Check exact text.
    *   Download errors (check console logs for permissions/404s).
    *   Unsupported image format errors (ensure Pillow is installed via `requirements.txt`; check conversion logs).
*   **`FileNotFoundError` for Template:** Check `TEMPLATE_FILENAME` in the report module and ensure the file exists in the `template/` folder.
