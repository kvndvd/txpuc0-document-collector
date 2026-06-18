import logging
import os
import time
from datetime import datetime, timedelta
from pathlib import Path

from searchCourt import search_court
from lndocData import lndoc_check
from duplicateCheck import duplicate_check
from sendEmail import sendEmail
from emailMsg import emailMsg
from timeRun import totalRun


LNDOC_DEFAULT_FROM_DATE = "11/01/2025"
COURT_LOOKBACK_DAYS = 7


def _safe_mkdir(folder: Path) -> None:
    """Create a folder and immediately verify that Windows can see it."""
    folder.mkdir(parents=True, exist_ok=True)
    if not folder.exists():
        raise FileNotFoundError(f"Folder was not created or is not visible to Windows: {folder}")

def setup_logger(log_file):
    """Set up logging after guaranteeing that the log folder exists.

    This is intentionally defensive because the app manager runs this collector
    from a subprocess. The log path must be absolute and its parent folder must
    exist before logging.FileHandler is created.
    """
    log_file = Path(log_file).resolve()
    _safe_mkdir(log_file.parent)

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    if logger.handlers:
        logger.handlers.clear()

    file_handler = logging.FileHandler(str(log_file), mode="w", encoding="utf-8")
    file_handler.setLevel(logging.INFO)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%m/%d/%Y %H:%M:%S",
    )

    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)


def create_run_paths():
    date_time = datetime.now().strftime("%m-%d-%Y_%I-%M-%S-%p")

    # Prefer an output folder provided by the app-manager wrapper.
    # Otherwise, use the current working directory.
    output_root_from_env = os.environ.get("TXPUC0_OUTPUT_ROOT", "").strip()
    if output_root_from_env:
        main_output_folder = Path(output_root_from_env).resolve()
    else:
        main_output_folder = (Path.cwd() / "STTXPUC0 Results").resolve()

    run_folder = main_output_folder / f"STTXPUC0_{date_time}"
    _safe_mkdir(run_folder)

    output_file = run_folder / f"STTXPUC0_Document_Collector_{date_time}.xlsx"
    log_file = run_folder / f"sttxpuc0_botrun_{date_time}.log"

    return run_folder, output_file, log_file


def get_auto_dates():
    """
    Court search:
        today - 7 days through today

    LN-DOC search:
        11/01/2025 through today
    """
    today = datetime.today()
    court_from_date = today - timedelta(days=COURT_LOOKBACK_DAYS)

    from_date_str = court_from_date.strftime("%m/%d/%Y")
    to_date_str = today.strftime("%m/%d/%Y")

    last_date = LNDOC_DEFAULT_FROM_DATE
    to_day_str = to_date_str

    return from_date_str, to_date_str, last_date, to_day_str


def main():
    startTime = time.time()

    output_folder, output_file, log_file = create_run_paths()
    setup_logger(log_file)

    logging.info("STTXPUC0 Document Collector started.")
    logging.info(f"Output folder: {output_folder}")
    logging.info(f"Excel file: {output_file}")
    logging.info(f"Log file: {log_file}")

    try:
        from_date_str, to_date_str, last_date, to_day_str = get_auto_dates()

        logging.info(f"Court Date Search: {from_date_str} to {to_date_str}")
        logging.info(f"LN-Doc Date Search: {last_date} to {to_day_str}")

        search_court(
            output_file=output_file,
            from_date_str=from_date_str,
            to_date_str=to_date_str,
        )

        lndoc_check(output_file, last_date, to_day_str)
        duplicate_count, include_count, download_failed_count = duplicate_check(output_file)

        endTime = time.time()
        emailBody = emailMsg(
            duplicate_count,
            include_count,
            download_failed_count,
            totalRun,
            endTime,
            startTime,
        )
        sendEmail(str(log_file), emailBody, from_date_str, to_date_str)

        logging.info("STTXPUC0 Document Collector completed successfully.")

    except Exception as error:
        logging.exception(f"STTXPUC0 Document Collector failed: {error}")
        raise


if __name__ == "__main__":
    main()
