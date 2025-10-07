import pandas as pd
import requests
import re
from bs4 import BeautifulSoup as bs
import os 
import glob
import time
import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import StaleElementReferenceException

N8N_WEBHOOK_URL = "https://anasellll.app.n8n.cloud/webhook/f234915f-8cdc-4838-8bf8-c3ee74680513"
DOWNLOAD_DIR = "/home/runner/work/downloads"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

USERNAME = "contact@targetupconsulting.com"
PASSWORD = "TargetUp2024@@"

# ------------------------
# Selenium setup
# ------------------------
options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
}
options.add_experimental_option("prefs", prefs)
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")

service = Service("/usr/bin/chromedriver")
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 15)

driver.get("https://www.offresonline.com/")

driver.find_element(By.CSS_SELECTOR,"#main-nav > ul > li:nth-child(2)").click() # se connecter
time.sleep(1)
driver.find_element(By.CSS_SELECTOR, "#Login").send_keys("TARGETUP") # User Name
time.sleep(1)
driver.find_element(By.CSS_SELECTOR, "#pwd").send_keys("TARGETUP2024") # Password
time.sleep(1)
driver.find_element(By.CSS_SELECTOR, "#buuuttt").click() # Cliker

driver.find_element(By.CSS_SELECTOR, "#ctl00_Linkf30").click() # AO de Jour
time.sleep(2)


def extract_popup_data(driver, wait):
    """
    This function extracts all information from the currently open popup.
    It assumes the driver has already switched to the correct iframe.
    """
    details = {}
    try:
        # --- Extract Title Info (Reference and Publication Date) ---
        title_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.span_tota")))
        title_text = title_element.text
        # Example: "Détails d'appel d'offre N° 23/MDA/2025 publié le 15/09/2025"
        parts = title_text.split('publié le')
        if len(parts) == 2:
            details['Référence N°'] = parts[0].replace("Détails d'appel d'offre N°", "").strip()
            details['Date de Publication'] = parts[1].strip()
        else:
            details['Référence N°'] = 'Not Found'
            details['Date de Publication'] = 'Not Found'

        # --- Extract Key-Value Pairs from the main table ---
        # Find all rows in the main details table
        detail_rows = driver.find_elements(By.XPATH, "//table[@width='800px']//tr")
        
        for row in detail_rows:
            try:
                # We look for rows that have exactly two columns (key and value)
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) == 2:
                    key = cells[0].text.strip().replace(':', '').replace('&nbsp;', '').strip()
                    value = cells[1].text.strip()
                    if key and value: # Only add if both key and value have content
                        details[key] = value
            except Exception:
                continue # Ignore rows that don't fit the pattern
        
        # --- Special case for the "Objet" which is in a nested table ---
        try:
            objet_element = driver.find_element(By.CSS_SELECTOR, "td.classltdtitreleftvueNBOBJ strong")
            details['Objet'] = objet_element.text.strip()
        except NoSuchElementException:
            details['Objet'] = None # Or a default value like 'Not Found'

        # --- Special case for "Date Limite" at the bottom ---
        try:
            date_limite_element = driver.find_element(By.XPATH, "//td[contains(text(), 'DATE LIMITE')]/following-sibling::td")
            details['DATE LIMITE'] = date_limite_element.text.strip()
        except NoSuchElementException:
            details['DATE LIMITE'] = None

    except TimeoutException:
        print("    - Timed out while waiting for popup content.")
        return None
    except Exception as e:
        print(f"    - An unexpected error occurred during extraction: {e}")
        return None
        
    return details


wait = WebDriverWait(driver, 10)
scraped_data = [] # This list will hold all our scraped dictionaries

try:
    # --- 3. LOCATE THE TABLE AND GET THE NUMBER OF ROWS ---
    print("Waiting for the main table to load...")
    table_body = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#tableao > tbody")))
    rows = table_body.find_elements(By.TAG_NAME, "tr")
    num_rows = len(rows)
    
    if num_rows == 0:
        print("No rows were found in the table. Exiting.")
    else:
        print(f"Found {num_rows} rows to process.")
        print("-" * 30)

    # --- 4. LOOP THROUGH EACH ROW ---
    for i in range(num_rows):
        print(f"Processing row {i + 1} of {num_rows}...")
        try:
            # Re-find the rows in each iteration to prevent StaleElementReferenceException
            current_rows = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#tableao > tbody"))).find_elements(By.TAG_NAME, "tr")
            current_rows[i].click()

            # --- Switch to the iframe that contains the popup ---
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "IframeEdit")))
            
            # --- Scrape the data ---
            print("  - Switched to iframe. Extracting data...")
            popup_info = extract_popup_data(driver, wait)
            if popup_info:
                scraped_data.append(popup_info)
                print("  - Data extracted successfully.")

            # --- IMPORTANT: Switch back to the main document ---
            driver.switch_to.default_content()
            
            # --- Close the popup ---
            close_button = wait.until(EC.element_to_be_clickable((By.ID, "btnCancel")))
            close_button.click()
            time.sleep(1) # Stable wait for popup to close
            
            print(f"Row {i + 1} processed.")
            print("-" * 30)

        except Exception as e:
            print(f"An error occurred on row {i + 1}: {e}")
            print("Refreshing page and trying to continue...")
            driver.refresh()
            time.sleep(3)
            continue

except TimeoutException:
    print("Error: The initial table was not found on the page.")
finally:
    # --- 5. CREATE DATAFRAME AND EXPORT TO CSV ---
    if scraped_data:
        print("\nCreating DataFrame from scraped data...")
        df = pd.DataFrame(scraped_data)
        
        # Reorder columns to have important ones first
        desired_order = ['Référence N°', 'Date de Publication', 'DATE LIMITE', 'ORGANISME', 'Objet', 'CAUTION', 'ESTIMATION FINANCIERE', 'VILLE(S)', 'CONTACT']
        # Get existing columns and add the rest, avoiding errors if a column doesn't exist
        existing_cols = [col for col in desired_order if col in df.columns]
        other_cols = [col for col in df.columns if col not in existing_cols]
        df = df[existing_cols + other_cols]
    else:
        print("\nNo data was scraped. No output file created.")

    # --- 6. CLEANUP ---
    print("Closing the browser.")

