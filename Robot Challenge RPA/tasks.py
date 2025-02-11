from robocorp import browser, vault, storage
from robocorp.tasks import task
from RPA.Excel.Files import Files as Excel

from pathlib import Path
import os
import requests

FILE_NAME = "challenge.xlsx"
OUTPUT_DIR = Path(os.environ.get("ROBOT_ARTIFACTS"))
EXCEL_URL = f"https://rpachallenge.com/assets/downloadFiles/{FILE_NAME}"
secrets = vault.get_secret("DemoVault")
locators = storage.get_json("Locators")


@task
def solve_challenge():
    """
    Solve the RPA challenge

    Downloads the source data excel and uses Playwright to solve rpachallenge.com from challenge
    """
    browser.configure(
        browser_engine="chromium",
        screenshot="only-on-failure",
        headless=False,
    )
    try:
        """
        vault = private data
        """
        excel_file = download_file(EXCEL_URL, OUTPUT_DIR, FILE_NAME)
        excel = Excel()
        excel.open_workbook(excel_file)
        rows = excel.read_worksheet_as_table("Sheet1", header=True)

        page = browser.goto(secrets["url"])
        page.click(locators["start"])
        for row in rows:
            fill_and_submit_form(row)
        element = page.locator("css=div.congratulations")
        browser.screenshot(element)

    finally:
        # Place for teardown and cleanups
        # Playwright handles browser closing
        print("Done")


def download_file(url: str, target_dir: Path, target_filename: str) -> str:
    """
    Downloads a file from the given url into the given folder with given filename.
    """
    target_dir.mkdir(exist_ok=True)
    response = requests.get(url)
    response.raise_for_status()  # This will raise an exception if the request fails
    local_filename = Path(target_dir, target_filename)
    with open(local_filename, "wb") as f:
        f.write(response.content)  # Write the content of the response to a file
    return local_filename


def fill_and_submit_form(row):
    """
    Fills a single form with the information of a single row in the Excel
    """
    page = browser.page()
    page.fill("//input[@ng-reflect-name='labelFirstName']", str(row["First Name"]))
    page.fill("//input[@ng-reflect-name='labelLastName']", str(row["Last Name"]))
    page.fill("//input[@ng-reflect-name='labelCompanyName']", str(row["Company Name"]))
    page.fill("//input[@ng-reflect-name='labelRole']", str(row["Role in Company"]))
    page.fill("//input[@ng-reflect-name='labelAddress']", str(row["Address"]))
    page.fill("//input[@ng-reflect-name='labelEmail']", str(row["Email"]))
    page.fill("//input[@ng-reflect-name='labelPhone']", str(row["Phone Number"]))
    page.click(locators["submit"])
