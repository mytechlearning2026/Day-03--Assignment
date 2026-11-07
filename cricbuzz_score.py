from playwright.sync_api import sync_playwright
import time

with sync_playwright() as p:

    browser = p.chromium.launch(
        channel="msedge",
        headless=False
    )

    page = browser.new_page(
        viewport={"width": 1366, "height": 768}
    )

    # Open Cricbuzz website
    page.goto(
        "https://www.cricbuzz.com/cricket-match/live-scores",
        wait_until="domcontentloaded"
    )

    # Wait for the page to display
    time.sleep(2)

    # Take screenshot
    page.screenshot(
        path="cricbuzz_score.png",
        full_page=True
    )

    print("Screenshot saved successfully!")

    # Keep browser open for observation
    time.sleep(2)

    browser.close()

import os
import time
import subprocess

from openpyxl import Workbook

# Create Excel workbook
workbook = Workbook()

sheet = workbook.active
sheet.title = "Live Cricket Scores"

# Add headings
sheet.append([
    "Match",
    "Team 1",
    "Team 1 Score",
    "Team 2",
    "Team 2 Score",
    "Match Status"
])

# Sample data
sheet.append([
    "India vs Australia",
    "India",
    "250/5",
    "Australia",
    "180/3",
    "Live"
])

# Save Excel file
file_path = os.path.abspath("cricbuzz_score.xlsx")

workbook.save(file_path)

print("Excel file saved successfully!")

# Open Excel file
excel_process = subprocess.Popen(
    ["cmd", "/c", "start", "", file_path],
    shell=False
)

print("Excel file opened successfully!")

# Wait for 2 seconds
time.sleep(3)

# Close Excel automatically
os.system("taskkill /IM EXCEL.EXE /F")

print("Excel closed successfully!")
