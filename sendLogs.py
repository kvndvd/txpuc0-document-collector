import logging
from pathlib import Path

import win32com.client as win32


def sendLogs(log_file, from_date_str, to_date_str):
    log_attachment = Path(log_file).resolve()

    if not log_attachment.exists():
        raise FileNotFoundError(f"Log file not found for email attachment: {log_attachment}")

    logging.info("Dispatching Outlook email for logs")
    logging.info(f"Email log attachment: {log_attachment}")

    outlook = win32.Dispatch("outlook.application")
    namespace = outlook.GetNamespace("MAPI")
    mail = outlook.CreateItem(0)
    mail.To = str(namespace.CurrentUser.Address)
    mail.Subject = f"TXPUC0 Document Collector {from_date_str}-{to_date_str}"
    mail.Attachments.Add(str(log_attachment))
    mail.Send()

    logging.info("Outlook log email sent successfully")
