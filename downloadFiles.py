import logging
import re
import time
from pathlib import Path

import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

PDF_HOST = "interchange.puc.texas.gov"
CHUNK_SIZE = 1024 * 1024
RETRY_STATUS_CODES = {500, 502, 503, 504}
MAX_DOWNLOAD_ATTEMPTS = 3


def get_download_filename(citation):
    """
    Return the formatted local PDF filename used for a downloaded included court file.

    Example:
        59711_9_1638911 -> LDC_SMD_59711_9_1638911_E2E.pdf
    """
    citation = normalize_text(citation)
    return f"LDC_SMD_{safe_filename(citation)}_E2E.pdf"


def download_included_file(output_file, citation, pdf_link):
    """
    Download one included TXPUC0 PDF into a Downloaded Files folder beside output_file.

    Returns True only when the PDF already exists or downloads successfully.
    Returns False when the download fails, so duplicateCheck.py can leave the
    original link in Excel and mark the row as a download failure.
    """
    citation = normalize_text(citation)
    pdf_link = normalize_text(pdf_link)

    if not citation:
        logging.warning("Download skipped: missing citation.")
        return False

    if not pdf_link:
        logging.warning(f"Download skipped for {citation}: missing PDF link.")
        return False

    output_file = Path(output_file)
    download_folder = output_file.parent / "Downloaded Files"
    download_folder.mkdir(parents=True, exist_ok=True)

    file_name = get_download_filename(citation)
    download_path = download_folder / file_name

    if download_path.exists() and download_path.stat().st_size > 0:
        logging.info(f"Download skipped for {citation}: file already exists.")
        return True

    return download_file(pdf_link, download_path, citation)


def download_file(url_download, file_save_path, include_filename):
    """
    Backward-compatible downloader based on the old working downloadFile.py.
    Uses verify=False because the TX PUC site can fail Python certificate validation.
    Retries temporary server errors such as HTTP 500/502/503/504.
    """
    url_download = normalize_text(url_download)
    file_save_path = Path(file_save_path)
    include_filename = normalize_text(include_filename)

    if not url_download:
        logging.error(f"{include_filename} - Failed to download. Missing URL.")
        return False

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/131.0.0.0 Safari/537.36"
        ),
        "Accept": "application/pdf,application/octet-stream,*/*;q=0.8",
        "Referer": "https://interchange.puc.texas.gov/Search/Daily",
        "Connection": "close",
    }

    try:
        file_save_path.parent.mkdir(parents=True, exist_ok=True)

        response = None
        for attempt in range(1, MAX_DOWNLOAD_ATTEMPTS + 1):
            response = requests.get(
                url_download,
                headers=headers,
                stream=True,
                verify=False,
                timeout=60,
            )

            if response.status_code == 200:
                break

            logging.warning(
                f"{include_filename} - Download attempt {attempt} failed. "
                f"HTTP Status: {response.status_code}"
            )

            if response.status_code in RETRY_STATUS_CODES and attempt < MAX_DOWNLOAD_ATTEMPTS:
                response.close()
                time.sleep(3)
                continue

            logging.error(
                f"{include_filename} - Failed to download. HTTP Status: {response.status_code}"
            )
            response.close()
            return False

        if response is None or response.status_code != 200:
            logging.error(f"{include_filename} - Failed to download. No successful response.")
            return False

        content_type = response.headers.get("Content-Type", "")
        total_size = int(response.headers.get("Content-Length", 0) or 0)

        logging.info(f"Downloading - {include_filename} - {total_size / 1024:.2f} KB")

        downloaded_size = 0
        first_chunk = True
        saw_pdf_signature = False

        with open(file_save_path, "wb") as file:
            for chunk in response.iter_content(chunk_size=CHUNK_SIZE):
                if not chunk:
                    continue

                if first_chunk:
                    saw_pdf_signature = chunk.lstrip().startswith(b"%PDF")
                    first_chunk = False

                file.write(chunk)
                downloaded_size += len(chunk)

        response.close()

        if downloaded_size == 0:
            logging.error(f"{include_filename} - Download failed. Empty file received.")
            delete_partial_file(file_save_path)
            return False

        if "pdf" not in content_type.lower() and not saw_pdf_signature:
            logging.error(
                f"{include_filename} - Invalid content type: {content_type}. "
                "Downloaded file does not start with PDF signature."
            )
            delete_partial_file(file_save_path)
            return False

        logging.info(
            f"Downloaded Successfully - {downloaded_size / 1024:.2f} KB"
        )

        if total_size > 0 and downloaded_size != total_size:
            logging.warning(
                f"{include_filename} - Warning: Expected {total_size / 1024:.2f} KB, "
                f"but downloaded {downloaded_size / 1024:.2f} KB"
            )

        return True

    except requests.exceptions.RequestException as error:
        logging.error(f"Error downloading {include_filename}: {error}")
        delete_partial_file(file_save_path)
        return False

    except Exception as error:
        logging.error(f"Unexpected error downloading {include_filename}: {error}")
        delete_partial_file(file_save_path)
        return False


def delete_partial_file(file_path):
    try:
        file_path = Path(file_path)
        if file_path.exists():
            file_path.unlink()
    except Exception as error:
        logging.warning(f"Could not delete partial download {file_path}: {error}")


def safe_filename(value):
    value = normalize_text(value)
    value = re.sub(r'[<>:"/\\|?*]', "_", value)
    value = re.sub(r"\s+", "_", value)
    return value.strip("._ ") or "download"


def normalize_text(value):
    if value is None:
        return ""

    return str(value).replace("\xa0", " ").strip()
