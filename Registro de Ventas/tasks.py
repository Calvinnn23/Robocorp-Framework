from robocorp.tasks import task
from robocorp import browser
import os
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF


@task
def robot_spare_bin():
    """
    Insert sales data and export as PDF
    """
    browser.configure(slowmo=100)
    open_web()
    log_in()
    download_file()
    fill_form_excel()
    capture_info()
    export_as_pdf()
    log_out()


def open_web():
    """
    Open url in browser
    """
    browser.goto("https://robotsparebinindustries.com/#/")


def log_in():
    """
    Log in web
    """
    page = browser.page()
    page.fill("#username", "maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")


def download_file():
    """
    Download excel file with sales data
    """
    http = HTTP()
    url = "https://robotsparebinindustries.com/SalesData.xlsx"
    destinyFolder = "data"
    os.makedirs(destinyFolder, exist_ok=True)
    http.download(url, destinyFolder, overwrite=True)


def fill_form(fileRows):
    """
    Fill form with excel file data
    """
    page = browser.page()
    page.fill("#firstname", str(fileRows["First Name"]))
    page.fill("#lastname", str(fileRows["Last Name"]))
    page.select_option("#salestarget", str(fileRows["Sales Target"]))
    page.fill("#salesresult", str(fileRows["Sales"]))
    page.click("button:text('Submit')")


def fill_form_excel():
    """
    Read each row from excel file as a table
    """
    excel = Files()
    excel.open_workbook("data/SalesData.xlsx")
    worksheet = excel.read_worksheet_as_table("data", header=True)
    excel.close_workbook()

    for row in worksheet:
        fill_form(row)


def capture_info():
    """
    Take a screenshot of the sales summary
    """
    page = browser.page()
    page.screenshot(path="screenshot/SalesSummary.png")


def export_as_pdf():
    """
    Save as pdf the sales detailed
    """
    page = browser.page()
    sales_html = page.locator("#sales-results").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(sales_html, "pdf/SalesResults.pdf")


def log_out():
    """
    Log out web
    """
    page = browser.page()
    page.click("#logout")
