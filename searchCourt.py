import logging
import re
import time
from datetime import datetime
from urllib.parse import parse_qs, urlparse, urljoin

from openpyxl import Workbook
from scrapling import Selector

from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait

from initialize import user_validation_popup
from session import session


COURT_START_URL = "https://interchange.puc.texas.gov/Search/Daily"

INCLUDE_TERMS = (
    "order no",
    "response to order no",
    "commission staff’s response to order",
    "commission staff's response to order",
)
EXCLUDE_TERMS = (
    "soah order no",
    "consent form",
    "additional information in response to order",
)

DEFAULT_DATE_FORMAT = "%m/%d/%Y"


def search_court(output_file, from_date_str=None, to_date_str=None):

    driver = session()
    wait = WebDriverWait(driver, 20)

    rows = []
    visited_pages = set()

    from_date_str, to_date_str = norm_date(from_date_str, to_date_str)

    try:
        logging.info(f"Opening STTXPUC0 court site: {COURT_START_URL}")

        driver.get(COURT_START_URL)
        pagewait_ready(driver, wait)
        user_validation(driver)

        daily_search(driver, wait, from_date_str, to_date_str)

        current_url = driver.current_url
        first_tab = driver.current_window_handle

        no_docs_text = get_no_docs_text(Selector(driver.page_source))
        logging.info(no_docs_text)

        if has_zero_documents(no_docs_text):
            logging.info("No document found")
            create_excel_file(rows, output_file)
            return

        while current_url and current_url not in visited_pages:
            visited_pages.add(current_url)

            pagewait_result(driver, wait)
            html = get_pagesource(driver, wait)
            page = Selector(html)

            scrapecurrent_page(page, driver, wait, first_tab, rows)

            next_url = get_next_page_url(page, driver.current_url)
            if not next_url:
                break

            logging.info(f"Next court page: {next_url}")
            driver.get(next_url)
            pagewait_ready(driver, wait)
            user_validation(driver)
            current_url = driver.current_url

        create_excel_file(rows, output_file)

        logging.info("Court check completed.")
        logging.info(f"Total court PDFs collected: {len(get_unique_rows(rows))}")

    except TimeoutException as error:
        logging.exception(f"Court check timed out: {error}")
        raise

    except Exception as error:
        logging.exception(f"Court check failed: {error}")
        raise

    finally:
        driver.quit()
        logging.info("Court browser closed.")

def getitemorder_nuber(document_url):
    if not document_url:
        return "", ""

    parsed_url = urlparse(document_url)
    query_values = parse_qs(parsed_url.query)

    order_number = query_values.get("controlNumber", [""])[0]
    item_number = query_values.get("itemNumber", [""])[0]

    return order_number, item_number


def norm_date(from_date_str=None, to_date_str=None):
    if from_date_str and to_date_str:
        return from_date_str, to_date_str

    today = datetime.now().strftime(DEFAULT_DATE_FORMAT)

    if not from_date_str:
        from_date_str = today
    if not to_date_str:
        to_date_str = today

    logging.info("No complete date range was passed to search_court; using today's date range.")
    return from_date_str, to_date_str


def pagewait_ready(driver, wait):
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")


def pagewait_result(driver, wait):
    wait.until(ec.presence_of_element_located((By.ID, "main")))


def user_validation(driver):
    while has_user_validation(driver):
        logging.info("User Validation required. Complete the Captcha.")

        if user_validation_popup(driver):
            break

        logging.warning("Captcha not completed. Refreshing and retrying...")
        driver.refresh()
        time.sleep(2)


def has_user_validation(driver):
    try:
        user_valid = driver.find_element(By.XPATH, "/html/body/h3").text
        return "User validation required to continue.." in user_valid
    except NoSuchElementException:
        return False


def daily_search(driver, wait, from_date_str, to_date_str):
    wait.until(ec.presence_of_element_located((By.ID, "DateFiledFrom")))

    from_date = driver.find_element(By.ID, "DateFiledFrom")
    to_date = driver.find_element(By.ID, "DateFiledTo")

    from_date.clear()
    to_date.clear()

    from_date.send_keys(from_date_str)
    to_date.send_keys(to_date_str)

    driver.find_element(By.ID, "searchButton").click()
    logging.info("Searching court site...")

    pagewait_result(driver, wait)


def get_pagesource(driver, wait):
    driver.switch_to.default_content()
    pagewait_ready(driver, wait)
    return driver.page_source


def scrapecurrent_page(page, driver, wait, first_tab, rows):
    result_rows = getresult_rows(page, driver.current_url)
    collected_before = len(rows)

    for result_row in result_rows:
        row_text = result_row["row_text"]

        if exclude_row(row_text):
            logging.info(f"Exclude row: {result_row['display_text']}")
            continue

        if not include_row(row_text):
            continue

        document_url = result_row["document_url"]
        if not document_url:
            logging.warning(f"Included row has no document link; skipping: {result_row['display_text']}")
            continue

        order_number = result_row["order_number"]
        item_number = result_row["item_number"]

        logging.info(f"Include row: {result_row['display_text']}")
        logging.info(f"Opening linked item {order_number}_{item_number}: {document_url}")

        scrape_page(driver, wait, first_tab, document_url, rows)

    logging.info(f"Total included on this page: {len(rows) - collected_before}")


def getresult_rows(page, base_url):
    parsed_rows = []

    for row in page.css("#main table tr"):
        cells = row.css("td")

        if len(cells) < 2:
            continue

        cell_texts = [normalize_text(cell.text) for cell in cells]
        item_link = get_first_link(cells[1], base_url)

        order_number = cell_texts[0] if len(cell_texts) > 0 else ""
        item_number = cell_texts[1] if len(cell_texts) > 1 else ""

        link_order_number, link_item_number = getitemorder_nuber(item_link)

        if not order_number:
            order_number = link_order_number

        if not item_number:
            item_number = link_item_number

        display_cell_texts = list(cell_texts)

        if len(display_cell_texts) > 0 and order_number:
            display_cell_texts[0] = order_number

        if len(display_cell_texts) > 1 and item_number:
            display_cell_texts[1] = item_number

        parsed_rows.append({
            "order_number": order_number,
            "item_number": item_number,
            "document_url": item_link,
            "row_text": " ".join(display_cell_texts).lower(),
            "display_text": " | ".join(display_cell_texts),
        })

    return parsed_rows


def exclude_row(row_text):
    return any(term in row_text for term in EXCLUDE_TERMS)


def include_row(row_text):
    return any(term in row_text for term in INCLUDE_TERMS)


def scrape_page(driver, wait, first_tab, document_url, rows):
    driver.execute_script(
        "window.open(arguments[0], '_blank', 'noopener,noreferrer');",
        document_url,
    )
    driver.switch_to.window(driver.window_handles[-1])

    try:
        pagewait_ready(driver, wait)
        user_validation(driver)
        pagewait_result(driver, wait)

        html = get_pagesource(driver, wait)
        page = Selector(html)
        pdf_entries = get_pdf_entries(page, driver.current_url)

        if not pdf_entries:
            logging.info("No PDF files found in this table.")
            return

        check_for_merge = len(pdf_entries) > 1

        for entry in pdf_entries:
            rows.append({
                "From Court": entry["name"],
                "From Ln-Doc": "",
                "Result": "Check for merge" if check_for_merge else "",
                "Link": entry["link"] or document_url,
            })

            logging.info(f"Collected court PDF: {entry['name']}")

    finally:
        if len(driver.window_handles) > 1:
            driver.close()
        driver.switch_to.window(first_tab)

def main_table(page):
    tables = page.css("#main table")

    if not tables:
        return None

    for table in tables:
        header_text = normalize_text(table.text).lower()

        if "type" in header_text and "name" in header_text:
            return table

    return tables[0]


def tableindex_column(table):
    column_indexes = {}

    headers = table.css("thead th")

    if not headers:
        headers = table.css("tr th")

    for index, header in enumerate(headers):
        header_text = normalize_text(header.text).lower()

        if header_text:
            column_indexes[header_text] = index

    return column_indexes


def get_first_link_node(node):
    links = node.css("a[href]")

    if not links:
        return None

    return links[0]


def get_pdf_entries(page, base_url):
    pdf_entries = []

    table = main_table(page)

    if not table:
        logging.warning("Document table not found.")
        return pdf_entries

    column_indexes = tableindex_column(table)

    name_index = column_indexes.get("name", 0)
    type_index = column_indexes.get("type")

    if type_index is None:
        logging.warning("Could not find Type column in document table.")
        return pdf_entries

    for row in table.css("tbody tr"):
        cells = row.css("td")

        if len(cells) <= type_index:
            continue

        file_type = normalize_text(cells[type_index].text).upper()

        if file_type != "PDF":
            continue

        name_cell = cells[name_index] if len(cells) > name_index else cells[0]

        link_node = get_first_link_node(name_cell) or get_first_link_node(row)

        if link_node:
            name = normalize_text(link_node.text)
            href = link_node.attrib.get("href", "")
            link = urljoin(base_url, href) if href else ""
        else:
            name = normalize_text(name_cell.text)
            link = ""

        if not name:
            continue

        pdf_entries.append({
            "name": name,
            "link": link,
        })

        logging.info(f"PDF found: {name} | Link: {link}")

    return pdf_entries


def get_first_link(node, base_url):
    links = node.css("a[href]")

    if not links:
        return ""

    href = links[0].attrib.get("href", "")

    if not href:
        return ""

    return urljoin(base_url, href)


def get_next_page_url(page, base_url):
    next_nodes = page.css(".PagedList-skipToNext")

    if not next_nodes:
        logging.info("No additional pages found.")
        return None

    next_node = next_nodes[0]
    class_value = next_node.attrib.get("class", "")
    if isinstance(class_value, (list, tuple, set)):
        next_class = " ".join(class_value)
    else:
        next_class = str(class_value)

    if "disabled" in next_class.lower():
        logging.info("No additional pages found.")
        return None

    links = next_node.css("a[href]")

    if not links:
        logging.info("No additional pages found.")
        return None

    href = links[0].attrib.get("href", "")

    if not href:
        logging.warning("Next court page link found but href is empty.")
        return None

    return urljoin(base_url, href)


def get_no_docs_text(page):
    count_nodes = page.css("#main p")

    if not count_nodes:
        return ""

    return normalize_text(count_nodes[0].text)


def has_zero_documents(no_docs_text):
    match = re.search(r"\d+", no_docs_text or "")
    return bool(match and int(match.group()) == 0)


def create_excel_file(rows, output_file):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Document Collector"

    headers = ["From Court", "From Ln-Doc", "Result", "Link"]
    sheet.append(headers)

    for row in get_unique_rows(rows):
        sheet.append([
            row["From Court"],
            row["From Ln-Doc"],
            row["Result"],
            row["Link"],
        ])

    workbook.save(output_file)


def get_unique_rows(rows):
    unique_rows = []
    seen = set()

    for row in rows:
        unique_key = (row["From Court"], row["Link"])

        if unique_key in seen:
            continue

        seen.add(unique_key)
        unique_rows.append(row)

    return unique_rows


def normalize_text(value):
    return " ".join(str(value).replace("\xa0", " ").split())


# Backward-compatible alias if another script still imports court_check.
court_check = search_court
