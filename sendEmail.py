import logging
from pathlib import Path

import win32com.client as win32


def sendEmail(log_file, emailBody, from_date_str, to_date_str):
    """
    Send the run summary email and attach the log file.

    Outlook COM requires Attachments.Add(Source) to receive a plain string path,
    not a pathlib.Path / WindowsPath object.
    """
    log_attachment = Path(log_file).resolve()

    if not log_attachment.exists():
        raise FileNotFoundError(f"Log file not found for email attachment: {log_attachment}")

    logging.info("Dispatching Outlook email for results")
    logging.info(f"Email log attachment: {log_attachment}")

    outlook = win32.Dispatch("outlook.application")
    namespace = outlook.GetNamespace("MAPI")

    recipient = get_current_user_email(namespace)

    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = f"TXPUC0 Document Collector {from_date_str}-{to_date_str}"
    mail.HTMLBody = emailBody
    mail.Attachments.Add(str(log_attachment))
    mail.Send()

    logging.info("Outlook email sent successfully")


def get_current_user_email(namespace):
    """Return the current Outlook user's SMTP/address string."""
    try:
        current_user = namespace.CurrentUser
        address = getattr(current_user, "Address", None)
        if address:
            return str(address)
        name = getattr(current_user, "Name", None)
        if name:
            return str(name)
        return str(current_user)
    except Exception:
        logging.warning("Could not resolve Outlook current user. Falling back to blank recipient.")
        return ""
