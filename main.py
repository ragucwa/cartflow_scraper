import time
from scraper import format_excel, scrape_products, write_to_excel
from utils import setup_logging


def main():
    setup_logging()
    df = scrape_products()
    write_to_excel(df)
    format_excel()


if __name__ == "__main__":
    start_time = time.time()

    main()

    end_time = time.time()
    print(f"Time taken: {end_time - start_time:.2f} seconds")
