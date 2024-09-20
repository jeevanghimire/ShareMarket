from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import pandas as pd
import time

# URL of the company list page
url = 'https://www.sharesansar.com/company-list'

# Setup Selenium WebDriver (for Chrome)
driver = webdriver.Chrome()  # Ensure you have the correct ChromeDriver

# Function to scrape data by sector
def scrape_sector_data(sector):
    driver.get(url)
    time.sleep(2)  # Allow the page to load

    # Wait for the sector dropdown to be clickable
    wait = WebDriverWait(driver, 10)
    select_element = wait.until(EC.element_to_be_clickable((By.ID, 'sector')))

    # Use Select to choose the dropdown option
    try:
        select = Select(select_element)
        select.select_by_visible_text(sector)
    except Exception as e:
        print(f"Skipping sector {sector}: {e}")
        return None  # Skip this sector if there's an issue

    # Click the search button after selecting the sector
    search_button = driver.find_element(By.XPATH,  "//button[contains(text(), 'Search')]")
    search_button.click()

    # Wait for the table to load after selecting the sector
    table = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'table.table-striped.table-hover')))
    
    headers = [th.text.strip() for th in table.find_elements(By.TAG_NAME, 'th')]
    rows = []
    for tr in table.find_elements(By.TAG_NAME, 'tr')[1:]:  # Skip the header row
        cells = [td.text.strip() for td in tr.find_elements(By.TAG_NAME, 'td')]
        rows.append(cells)

    return pd.DataFrame(rows, columns=headers)

# List of sectors to scrape
sectors = [
    'Commercial Bank', 'Corporate Debentures', 'Development Bank',
    'Finance', 'Government Bonds', 'Hotel & Tourism', 'Hydropower', 'Investment', 'Life Insurance', 'Manufacturing and Processing',
    'Microfinance', 'Mutual Fund', 'Non-Life Insurance', 'Others',  'Preference Share', 'Promoter Share',  'Trading'
]

# Create an Excel writer object
excel_file = 'stock_data_by_sector_selenium.xlsx'
writer = pd.ExcelWriter(excel_file, engine='openpyxl')

# Loop through each sector and create a sheet
for sector in sectors:
    df = scrape_sector_data(sector)
    if df is not None:
        df.to_excel(writer, sheet_name=sector, index=False)

# Close the Excel writer
writer.close()
print(f"Data has been written to {excel_file}")

driver.quit()