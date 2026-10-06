
"""
I know basic Python and PyAutoGUI, but I’m not a developer. Please write a complete, beginner-friendly Python script that:

1. Opens Microsoft Edge and visits Yahoo Finance.
2. Gets the current S&P 500 index value from Yahoo Finance without needing me to attach a screenshot.
3. Takes and saves a screenshot of the browser.
4. Opens Excel and creates a worksheet with four columns: Date & Time, Stock Name, Price, and Comment. Enter the current date and time, “S&P 500” as the stock name, the fetched index value as the price, and “Today activity success” as the comment.
5. Saves the workbook as `daily_report_YYYY-MM-DD.xlsx` in the current folder. If a report with that name already exists, replace it. If that report is currently open in Excel, close it first.
6. Saves a screenshot of the completed worksheet, closes Excel, and opens the report folder in File Explorer.

Choose a straightforward way to do this. Before giving me the code, list any extra Python packages or software I need, explain in plain language what each is for, and give me the exact commands or steps to install them. Don’t assume I know package names or technical terms. Then provide the full script and simple instructions for running it. Include helpful messages if Yahoo Finance doesn’t load or the price can’t be found."""
""""""



import time
from datetime import datetime
from pathlib import Path
import subprocess
import sys

import pyautogui
import yfinance as yf
import win32com.client


pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1

# Save the files in the same folder as this script.
output_folder = Path(__file__).resolve().parent
today = datetime.now()
date_stamp = today.strftime("%Y-%m-%d")

browser_screenshot = output_folder / f"yahoo_finance_{date_stamp}.png"
excel_file = output_folder / f"daily_report_{date_stamp}.xlsx"
excel_screenshot = output_folder / f"daily_report_{date_stamp}.png"


def get_latest_sp500_price():
    """Fetch the latest available S&P 500 value from Yahoo Finance."""
    ticker = yf.Ticker("^GSPC")

    # Try the latest minute data first.
    prices = ticker.history(period="1d", interval="1m")

    # If minute data isn't available, use the latest daily value.
    if prices.empty or prices["Close"].dropna().empty:
        prices = ticker.history(period="5d", interval="1d")

    if prices.empty or prices["Close"].dropna().empty:
        raise ValueError("Yahoo Finance did not return an S&P 500 price.")

    return float(prices["Close"].dropna().iloc[-1])


def close_report_if_open(report_path):
    """Close this report in Excel if it is already open."""
    try:
        excel_app = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        # Excel isn't currently running.
        return

    target_path = str(report_path.resolve()).casefold()

    # Go backward because closing a workbook changes the workbook list.
    for index in range(excel_app.Workbooks.Count, 0, -1):
        workbook = excel_app.Workbooks(index)

        try:
            open_path = str(Path(workbook.FullName).resolve()).casefold()

            if open_path == target_path:
                # Discard the old copy because the report will be replaced.
                workbook.Close(SaveChanges=False)
                print("Closed the report that was already open in Excel.")
        except Exception:
            continue


def take_browser_screenshot():
    """Open the S&P 500 Yahoo Finance page and save a screenshot."""
    print("Opening Yahoo Finance in Microsoft Edge...")

    pyautogui.hotkey("win", "r")
    time.sleep(1)
    pyautogui.write(
        "msedge https://finance.yahoo.com/quote/%5EGSPC/",
        interval=0.03
    )
    pyautogui.press("enter")

    # Give Edge and the Yahoo Finance page time to load.
    time.sleep(15)

    pyautogui.screenshot().save(browser_screenshot)
    print(f"Browser screenshot saved: {browser_screenshot}")

    # Close the Edge window opened for this report.
    pyautogui.hotkey("alt", "f4")
    time.sleep(2)


def create_excel_report(price):
    """Create, save, and screenshot the Excel report."""
    close_report_if_open(excel_file)

    # Remove the old report file, if one exists.
    if excel_file.exists():
        try:
            excel_file.unlink()
        except PermissionError:
            raise PermissionError(
                "The report file is still in use. Close it in Excel and run "
                "the script again."
            )

    excel_app = None
    workbook = None

    try:
        # Start a separate Excel window for the new report.
        excel_app = win32com.client.DispatchEx("Excel.Application")
        excel_app.Visible = True
        excel_app.DisplayAlerts = False
        excel_app.WindowState = -4137  # Maximize Excel.

        workbook = excel_app.Workbooks.Add()
        sheet = workbook.Worksheets(1)
        sheet.Name = "Daily Report"

        # Add the headers and the report row.
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        sheet.Range("A1:D2").Value = (
            ("Date & Time", "Stock Name", "Price", "Comment"),
            (current_time, "S&P 500", price, "Today activity success")
        )

        # Format the header and price, and fit the columns to their contents.
        sheet.Range("A1:D1").Font.Bold = True
        sheet.Range("C2").NumberFormat = "#,##0.00"
        sheet.Columns("A:D").AutoFit()

        # Save the workbook as an Excel file.
        workbook.SaveAs(str(excel_file), FileFormat=51)
        sheet.Activate()
        sheet.Range("A1").Select()
        excel_app.Activate()

        # Give Excel time to display the finished sheet before taking its image.
        time.sleep(3)
        pyautogui.screenshot().save(excel_screenshot)

        print(f"Excel report saved: {excel_file}")
        print(f"Excel screenshot saved: {excel_screenshot}")

    finally:
        # Close the workbook and Excel even if saving the screenshot fails.
        if workbook is not None:
            workbook.Close(SaveChanges=True)

        if excel_app is not None:
            excel_app.Quit()


def main():
    try:
        take_browser_screenshot()

        print("Fetching the latest available S&P 500 value...")
        price = get_latest_sp500_price()
        print(f"S&P 500 value: {price:,.2f}")

        create_excel_report(price)

        # Open the folder containing the report and screenshots.
        subprocess.Popen(["explorer", str(output_folder)])
        print(f"Opened report folder: {output_folder}")

    except Exception as error:
        print(f"\nThe report could not be completed: {error}")
        print(
            "Check your internet connection, confirm Edge and Excel are "
            "installed, and make sure the report isn't open elsewhere."
        )
        sys.exit(1)


if __name__ == "__main__":
    main()