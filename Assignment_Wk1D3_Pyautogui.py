
"""
Yahoo Finance Screenshot and Daily Excel Report Automation
"""

import time
from datetime import datetime
from pathlib import Path

import pyautogui
import win32com.client

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1


def close_report_if_open(report_path):
    """Close the matching report workbook, then remove the old file."""
    try:
        excel_app = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        excel_app = None  # Excel is not already running.

    if excel_app is not None:
        target_path = str(report_path.resolve()).casefold()

        # Iterate backward because closing a workbook changes the collection.
        for index in range(excel_app.Workbooks.Count, 0, -1):
            workbook = excel_app.Workbooks(index)
            try:
                workbook_path = str(Path(workbook.FullName).resolve()).casefold()
                if workbook_path == target_path:
                    workbook.Close(SaveChanges=False)
            except Exception:
                pass

    # Removing the old file prevents an overwrite confirmation during Save As.
    if report_path.exists():
        report_path.unlink()


output_folder = Path.cwd()
today = datetime.now()
date_stamp = today.strftime("%Y-%m-%d")

browser_screenshot = output_folder / "screenshot.png"
excel_file = output_folder / f"daily_report_{date_stamp}.xlsx"
excel_screenshot = output_folder / f"daily_report_{date_stamp}.png"

# Close the existing report workbook and prepare to replace the file.
close_report_if_open(excel_file)

# Part 1: Open Yahoo Finance and take a screenshot.
print("Step 1: Open the browser")

pyautogui.hotkey("win", "r")
time.sleep(1)
pyautogui.write("msedge")
pyautogui.press("enter")
time.sleep(5)

pyautogui.hotkey("ctrl", "t")
pyautogui.write("https://finance.yahoo.com/")
pyautogui.press("enter")
time.sleep(10)  # Allow the page to load.

pyautogui.screenshot().save(browser_screenshot)
print(f"Browser screenshot saved: {browser_screenshot}")

pyautogui.hotkey("alt", "f4")
time.sleep(2)

# Part 2: Open Excel and enter the report.
print("Step 2: Open Excel")

pyautogui.hotkey("win", "r")
time.sleep(1)
pyautogui.write("excel")
pyautogui.press("enter")
time.sleep(10)  # Allow Excel to start.

pyautogui.hotkey("ctrl", "n")
time.sleep(3)

# Enter the four headers in row 1.
pyautogui.write("Date & Time")
pyautogui.press("tab")
pyautogui.write("Stock Name")
pyautogui.press("tab")
pyautogui.write("Price")
pyautogui.press("tab")
pyautogui.write("Comment")
pyautogui.press("enter")

# Enter the report data in row 2.
pyautogui.write(today.strftime("%Y-%m-%d %H:%M:%S"))
pyautogui.press("tab")
pyautogui.write("S&P Global")
pyautogui.press("tab")
pyautogui.write("7,818.93")
pyautogui.press("tab")
pyautogui.write("Today activity success")
pyautogui.press("enter")

# Auto-fit columns so the text is visible.
pyautogui.hotkey("ctrl", "a")
pyautogui.hotkey("alt", "h")
pyautogui.press("o")
pyautogui.press("i")
time.sleep(1)

# Save the new workbook.
pyautogui.press("f12")
time.sleep(2)
pyautogui.hotkey("ctrl", "a")
pyautogui.write(str(excel_file), interval=0.02)
pyautogui.press("enter")
time.sleep(4)

# Save a screenshot of the completed Excel sheet.
pyautogui.screenshot().save(excel_screenshot)
print(f"Excel file saved: {excel_file}")
print(f"Excel screenshot saved: {excel_screenshot}")

# Close Excel, then open the folder containing the report.
pyautogui.hotkey("alt", "f4")
time.sleep(2)

pyautogui.hotkey("win", "r")
time.sleep(1)
pyautogui.write(f'explorer "{output_folder}"', interval=0.02)
pyautogui.press("enter")

print(f"Opened report folder: {output_folder}")