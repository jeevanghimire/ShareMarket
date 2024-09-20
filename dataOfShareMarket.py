from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
import pandas as pd
import time
import os

# URL of the company list page
url = 'https://www.sharesansar.com/company-list'

# Set up Selenium WebDriver (for Chrome)
driver = webdriver.Chrome()

# Function to wait until all rows are loaded
def wait_for_rows_to_load(tbody, wait):
    previous_row_count = 0
    while True:
        try:
            rows = tbody.find_elements(By.TAG_NAME, 'tr')
            current_row_count = len(rows)
            if current_row_count == previous_row_count:
                break  # No new rows loaded
            previous_row_count = current_row_count
            time.sleep(1)  # Wait a moment before checking again
        except StaleElementReferenceException:
            tbody = wait.until(EC.presence_of_element_located((By.TAG_NAME, 'tbody')))

# Function to scrape data by sector with pagination
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

    # Click the search button
    try:
        search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Search')]")))
        driver.execute_script("arguments[0].click();", search_button)
    except Exception as e:
        print(f"Error clicking the search button: {e}")
        return None

    all_rows = []
    while True:
        # Wait for the table to load after clicking search
        try:
            table = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'table-striped')))
            tbody = table.find_element(By.TAG_NAME, 'tbody')

            # Wait for all rows to load
            wait_for_rows_to_load(tbody, wait)

            # Extract table headers and rows
            headers = [th.text.strip() for th in table.find_elements(By.TAG_NAME, 'th')]

            # Extract table rows
            for tr in tbody.find_elements(By.TAG_NAME, 'tr'):
                try:
                    cells = [td.text.strip() for td in tr.find_elements(By.TAG_NAME, 'td')]
                    if cells:
                        all_rows.append(cells)
                except StaleElementReferenceException:
                    print("Stale element detected. Retrying...")
                    tbody = wait.until(EC.presence_of_element_located((By.TAG_NAME, 'tbody')))  # Reload the tbody element

            # Check for "Next" button and click if present
            try:
                next_button = driver.find_element(By.XPATH, "//a[contains(@class, 'paginate_button next')]")
                if 'disabled' in next_button.get_attribute('class'):
                    break  # If "Next" is disabled, break the loop

                next_button.click()  # Click the "Next" button
                time.sleep(2)  # Allow time for the next page to load
            except Exception:
                break  # No "Next" button, end pagination

        except TimeoutException as e:
            print(f"Timeout while loading data: {e}")
            return None

    return pd.DataFrame(all_rows, columns=headers)

# List of sectors to scrape
sectors = [
    'Commercial Bank', 'Corporate Debentures', 'Development Bank',
    'Finance', 'Government Bonds', 'Hotel & Tourism', 'Hydropower', 'Investment',
    'Life Insurance', 'Manufacturing and Processing', 'Microfinance',
    'Mutual Fund', 'Non-Life Insurance', 'Others', 'Preference Share', 'Promoter Share', 'Trading'
]

# Create a folder for the Excel files
download_dir = r"./ShareMarket/Download"
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

# Loop through each sector and save its data into separate Excel files
for sector in sectors:
    df = scrape_sector_data(sector)
    if df is not None and not df.empty:  # Check if the DataFrame is not empty
        excel_file = os.path.join(download_dir, f'{sector}_data.xlsx')
        df.to_excel(excel_file, index=False)
        print(f"Data for sector {sector} has been written to {excel_file}")
    else:
        print(f"No data found for sector {sector}.")

# Close the driver
driver.quit()