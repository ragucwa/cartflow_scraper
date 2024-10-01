from openpyxl import load_workbook
import pandas as pd
from playwright.sync_api import TimeoutError, Page

from browser_context import BrowserContextManager
from locators import (
    NAV_BUTTON,
    LIST_OF_CATEGORIES,
    SEE_ALL_BUTTON,
    SUBCATEGORIES,
    PRODUCT_TITLE,
    NEXT_PAGE_BUTTON,
    CLEAR_BUTTON,
)
from constants import MAIN_URL, OUTPUT_FILE_NAME
import logging


def get_main_categories(page: Page):
    page.click(NAV_BUTTON)
    return page.locator(LIST_OF_CATEGORIES).all_inner_texts()


def get_subcategories(page: Page):
    try:
        page.locator(SEE_ALL_BUTTON).click(timeout=5000)
    except TimeoutError as e:
        logging.error(f"TimeoutError occurred: {e}")

    return page.locator(SUBCATEGORIES).all_inner_texts()


def get_products_titles(page: Page):
    titles = []
    while True:
        try:
            page.locator(PRODUCT_TITLE).first.wait_for()
            titles.extend(page.locator(PRODUCT_TITLE).all_inner_texts())
            page.locator(NEXT_PAGE_BUTTON).click(timeout=5000)
        except TimeoutError as e:
            logging.error(f"TimeoutError occurred: {e}")
            break
    return titles


def scrape_products():
    page = BrowserContextManager(browser_type="chromium")
    df = pd.DataFrame(columns=["Category", "Subcategory", "Product"])

    with page as page:
        page.goto(MAIN_URL)
        logging.info(f"Page title: {page.title()}")
        main_categories = get_main_categories(page)

        for x, category in enumerate(main_categories):
            page.locator(LIST_OF_CATEGORIES).nth(x).click()
            subcategories = get_subcategories(page)

            for y, subcategory in enumerate(subcategories):
                page.locator(SUBCATEGORIES).nth(y).click()
                titles = get_products_titles(page)

                for title in titles:
                    single_product = {
                        "Category": category,
                        "Subcategory": subcategory,
                        "Product": title,
                    }
                    df = pd.concat(
                        [df, pd.DataFrame([single_product])], ignore_index=True
                    )

                page.locator(SUBCATEGORIES).nth(y).click()

            page.locator(CLEAR_BUTTON).click()
            page.reload()
            page.click(NAV_BUTTON)

    return df


def write_to_excel(df: pd.DataFrame):
    with pd.ExcelWriter(OUTPUT_FILE_NAME) as writer:
        for category in df["Category"].unique():
            df[df["Category"] == category].to_excel(
                writer, sheet_name=category, index=False, engine="openpyxl"
            )


def format_excel():
    workbook = load_workbook(OUTPUT_FILE_NAME)
    sheet = workbook.active

    for sheet in workbook.worksheets:
        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except Exception as e:
                    logging.error(f"An error occurred: {e}")
                    pass
            adjusted_width = max_length + 2
            sheet.column_dimensions[column_letter].width = adjusted_width

        sheet.auto_filter.ref = sheet.dimensions

    workbook.save(OUTPUT_FILE_NAME)
