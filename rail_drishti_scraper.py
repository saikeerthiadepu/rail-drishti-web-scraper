# ============================================
# Rail Drishti Web Scraping Project
# Developed by: Adepu Saikeerthi
# Year: 2023
# ============================================

# Purpose:
# Scrape freight revenue data from the Rail Drishti
# website and export Commodity-wise and Zone-wise
# data into an Excel workbook.

# Libraries used:
# Selenium       - browser automation
# BeautifulSoup  - HTML parsing
# Pandas         - data processing and Excel export


from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd


# --------------------------------------------
# 1. Configure Chrome WebDriver
# --------------------------------------------

chrome_options = Options()

# Run Chrome in background
chrome_options.add_argument("--headless")

driver = webdriver.Chrome(options=chrome_options)


# --------------------------------------------
# 2. Open Rail Drishti webpage
# --------------------------------------------

url = "https://raildrishti.indianrailways.gov.in/raildrishti/raildrishtiv3/innerPages/FOISEarning.jsp"

driver.get(url)


# --------------------------------------------
# 3. Select required Month-to-Date category
# --------------------------------------------

try:

    button_xpath = "//*[@id='datawise']/div[3]/div"

    button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(
            (By.XPATH, button_xpath)
        )
    )

    button.click()

except Exception as error:

    print("Button could not be clicked:", error)


# --------------------------------------------
# 4. Get webpage source
# --------------------------------------------

page_source = driver.page_source

driver.quit()


# --------------------------------------------
# 5. Parse webpage using BeautifulSoup
# --------------------------------------------

soup = BeautifulSoup(
    page_source,
    "html.parser"
)


# --------------------------------------------
# 6. Extract Commodity-wise data
# --------------------------------------------

table1 = soup.find("table")

data_table1 = []

if table1:

    rows_table1 = table1.find_all("tr")

    for row in rows_table1:

        cells = row.find_all(["th", "td"])

        row_data = [
            cell.get_text(strip=True)
            for cell in cells
        ]

        data_table1.append(row_data)


# --------------------------------------------
# 7. Extract Zone-wise data
# --------------------------------------------

zone_tables = soup.find_all(
    "table",
    class_="table-condensed"
)

data_table2 = []

if len(zone_tables) > 1:

    zone_table = zone_tables[1]

    tbody_zone = zone_table.find("tbody")

    if tbody_zone:

        rows_table2 = tbody_zone.find_all("tr")

        for row in rows_table2:

            cells = row.find_all(["th", "td"])

            row_data = [
                cell.get_text(strip=True)
                for cell in cells
            ]

            data_table2.append(row_data)


# --------------------------------------------
# 8. Create Excel output
# --------------------------------------------

output_file = "Rail_Drishti_Scraped_Data.xlsx"


with pd.ExcelWriter(output_file) as writer:

    # Commodity-wise data

    if data_table1:

        df_table1 = pd.DataFrame(
            data_table1,
            columns=[
                "Commodity",
                "Current Year",
                "Previous Year",
                "Change %"
            ]
        )

        df_table1.to_excel(
            writer,
            sheet_name="COMMODITY WISE",
            index=False
        )


    # Zone-wise data

    if data_table2:

        df_table2 = pd.DataFrame(
            data_table2,
            columns=[
                "Zone",
                "Current Year",
                "Previous Year",
                "Change %"
            ]
        )

        df_table2.to_excel(
            writer,
            sheet_name="ZONE WISE",
            index=False
        )


print(
    f"Data saved successfully to {output_file}"
)
