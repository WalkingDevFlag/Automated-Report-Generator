# email_sender.py
"""
Handles the sending of generated reports via email using SMTP.

Connects to an SMTP server specified in config.py, authenticates,
constructs an email message with the report attached, and sends it
to the recipient specified in the Google Form response.
"""
from __future__ import annotations
import smtplib
import ssl
import os
import traceback
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Project Modules
import config # Import config to get email credentials and server info

# ===========================================================
# --- Email Content Configuration ---
# ===========================================================

# Define subject and body templates using placeholders
# These will be formatted with data from the specific report row.
DEFAULT_SUBJECT_TEMPLATE = "Your Generated Report: {template_choice}"
DEFAULT_BODY_TEMPLATE = """
Hello {submitter_name},

Please find your generated '{template_choice}' report attached.

This report was generated based on your submission at: {timestamp}

Regards,
The Automated Report System
"""
# You can customize the signature above

# ===========================================================
# --- End Email Content Configuration ---
# ===========================================================


def send_email_with_attachment(recipient_email: str, submitter_name: str, template_choice: str, timestamp: str, attachment_path: str) -> bool:
    """
    Sends an email with the generated report as an attachment.

    Args:
        recipient_email (str): The email address to send the report to.
        submitter_name (str): The name of the person who submitted the form.
        template_choice (str): The name of the template chosen by the user.
        timestamp (str): The timestamp of the form submission.
        attachment_path (str): The local file path to the generated report (.docx).

    Returns:
        bool: True if the email was sent successfully, False otherwise.
    """
    # --- Input Validation ---
    if not recipient_email or '@' not in recipient_email:
        print(f"[ERROR] Invalid recipient email address provided: '{recipient_email}'")
        return False
    if not config.EMAIL_SENDER or not config.EMAIL_PASSWORD:
         print("[ERROR] Email sender credentials (EMAIL_SENDER, EMAIL_PASSWORD) not configured in .env / config.py.")
         return False
    if not attachment_path or not os.path.exists(attachment_path):
        print(f"[ERROR] Email attachment file not found: {attachment_path}")
        return False

    print(f"[INFO] Preparing email for recipient: {recipient_email}")

    # --- Create Email Message ---
    try:
        message = MIMEMultipart()
        message["From"] = config.EMAIL_SENDER
        message["To"] = recipient_email
        # Format subject and body using templates and provided arguments
        message["Subject"] = DEFAULT_SUBJECT_TEMPLATE.format(template_choice=template_choice)

        body = DEFAULT_BODY_TEMPLATE.format(
            submitter_name=submitter_name if submitter_name else "User", # Use a fallback if name is empty
            template_choice=template_choice,
            timestamp=timestamp
        )
        message.attach(MIMEText(body, "plain"))
        print("[INFO] Email body composed.")
    except Exception as e:
        print(f"[ERROR] Failed to create email message structure: {e}")
        return False

    # --- Prepare Attachment ---
    filename = os.path.basename(attachment_path)
    try:
        with open(attachment_path, "rb") as attachment_file:
            # Use MIMEBase for generic binary attachment (suitable for .docx)
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment_file.read())
        # Encode file in base64
        encoders.encode_base64(part)
        # Add header to signify attachment
        part.add_header("Content-Disposition", f"attachment; filename= {filename}")
        message.attach(part)
        print(f"[INFO] Attached file: {filename}")
    except FileNotFoundError:
        # This case should have been caught earlier, but double-check.
        print(f"[ERROR] Attachment file disappeared before attaching: {attachment_path}")
        return False
    except Exception as e:
        print(f"[ERROR] Failed to attach file '{filename}': {e}")
        traceback.print_exc()
        return False

    # --- Send the Email ---
    server = None # Initialize server variable outside try block
    try:
        # Create secure SSL context
        context = ssl.create_default_context()
        print(f"[INFO] Connecting to SMTP server {config.SMTP_SERVER} on port {config.SMTP_PORT}...")

        if config.EMAIL_USE_SSL: # Port 465 typically uses SSL from the start
            server = smtplib.SMTP_SSL(config.SMTP_SERVER, config.SMTP_PORT, context=context)
            print("[INFO] Connected via SSL.")
        else: # Port 587 typically starts insecure and upgrades to TLS
            server = smtplib.SMTP(config.SMTP_SERVER, config.SMTP_PORT)
            print("[INFO] Connected (pre-TLS). Starting TLS...")
            server.starttls(context=context)
            print("[INFO] TLS connection established.")

        # Login to the email account
        print(f"[INFO] Logging in as {config.EMAIL_SENDER}...")
        server.login(config.EMAIL_SENDER, config.EMAIL_PASSWORD)
        print("[INFO] SMTP login successful.")

        # Send the email
        print(f"[INFO] Sending email to {recipient_email}...")
        server.sendmail(config.EMAIL_SENDER, recipient_email, message.as_string())
        print(f"[SUCCESS] Email sent successfully to {recipient_email}.")
        return True

    except smtplib.SMTPAuthenticationError:
        print(f"[ERROR] SMTP Authentication failed for {config.EMAIL_SENDER}.")
        print("  >>> Suggestion: Verify EMAIL_SENDER and EMAIL_PASSWORD in .env.")
        print("  >>> If using Gmail, ensure you are using an App Password if 2FA is enabled.")
        return False
    except smtplib.SMTPConnectError as e:
         print(f"[ERROR] Could not connect to SMTP server {config.SMTP_SERVER}:{config.SMTP_PORT}. Error: {e}")
         print("  >>> Suggestion: Verify SMTP_SERVER and SMTP_PORT in config.py. Check firewall/network.")
         return False
    except smtplib.SMTPServerDisconnected:
        print("[ERROR] SMTP server disconnected unexpectedly.")
        print("  >>> Suggestion: This might be a temporary server issue or network problem.")
        return False
    except ssl.SSLError as e:
        print(f"[ERROR] SSL Error occurred: {e}")
        print(f"  >>> Suggestion: Verify SMTP port ({config.SMTP_PORT}) matches SSL/TLS requirement.")
        print("  >>> Check if the server certificate is valid.")
        return False
    except Exception as e:
        # Catch any other unexpected errors during sending
        print(f"[ERROR] An unexpected error occurred sending email: {e}")
        traceback.print_exc()
        return False
    finally:
        # Ensure the server connection is closed gracefully
        if server:
            try:
                print("[INFO] Closing SMTP connection.")
                server.quit()
            except Exception as quit_e:
                 print(f"[WARNING] Error closing SMTP connection: {quit_e}")