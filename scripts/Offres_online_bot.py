import os
import time
import re
import glob
import base64
import pandas as pd
import requests
import zipfile
import fitz  # PyMuPDF
from pdf2image import convert_from_path
from PIL import Image
import pytesseract
import docx
import shutil
from bs4 import BeautifulSoup as bs

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.options import Options

from openai import OpenAI

print("--- SCRIPT STARTED ---")

# --- Credentials and Configuration ---
OFFRES_USERNAME = "TARGETUP"
OFFRES_PASSWORD = "TARGETUP2024"
OPENAI_API_KEY = "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"
N8N_WEBHOOK_URL = "https://targetup.app.n8n.cloud/webhook/dc4cf7c8-b44e-4404-830d-ef7cf3e7b6ca"

print("Credentials and configuration loaded.")

# --- Setup download folder ---
workspace_path = os.getcwd()
download_folder = os.path.join(workspace_path, "downloads")
# Clean the downloads folder before starting
if os.path.exists(download_folder):
    shutil.rmtree(download_folder)
os.makedirs(download_folder, exist_ok=True)
print(f"Download folder cleaned and created at: {download_folder}")

# --- Chrome options ---
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_folder,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

# --- Helper Functions ---

def wait_for_download_complete(directory, timeout=90):
    """Waits for a new download to complete in the specified directory."""
    seconds = 0
    while seconds < timeout:
        if any(f.endswith('.crdownload') for f in os.listdir(directory)):
            time.sleep(1)
            seconds += 1
        else:
            time.sleep(2) # Extra wait to ensure file is fully written
            print("  Download appears complete.")
            return True
    print(f"  ⚠️ DOWNLOAD TIMEOUT after {timeout} seconds.")
    return False

def extract_text_from_pdf(file_path):
    text = ""
    try:
        with fitz.open(file_path) as doc:
            for page in doc:
                text += page.get_text("text") + "\n"
        if len(text.strip()) < 150:
            print("    -> Short PDF text found, falling back to OCR.")
            raise Exception("Short text, try OCR")
        return text.strip()
    except Exception:
        try:
            pages = convert_from_path(file_path)
            ocr_text = ""
            for p_img in pages:
                ocr_text += pytesseract.image_to_string(p_img, lang="fra+ara") + "\n"
            return ocr_text.strip()
        except Exception as e:
            print(f"    -> FATAL OCR ERROR for {os.path.basename(file_path)}: {e}")
            return ""

def extract_text_from_docx(file_path):
    try:
        doc = docx.Document(file_path)
        return "\n".join([p.text for p in doc.paragraphs if p.text])
    except Exception as e:
        print(f"    -> ERROR reading DOCX {os.path.basename(file_path)}: {e}")
        return ""

def extract_text_from_csv(file_path):
    try:
        df_csv = pd.read_csv(file_path, sep=None, engine='python', on_bad_lines='skip')
        return df_csv.to_string()
    except Exception as e:
        print(f"    -> ERROR reading CSV {os.path.basename(file_path)}: {e}")
        return ""

def process_file_for_text(file_path):
    """Processes a single file and returns its text content."""
    ext = os.path.splitext(file_path)[1].lower()
    print(f"  -> Reading file: {os.path.basename(file_path)}")
    if ext == ".pdf": return extract_text_from_pdf(file_path)
    elif ext in [".docx", ".doc"]: return extract_text_from_docx(file_path)
    elif ext == ".csv": return extract_text_from_csv(file_path)
    return ""

def cleanup_files(paths_to_delete):
    """Safely removes files and directories."""
    for path in paths_to_delete:
        try:
            if os.path.isfile(path):
                os.remove(path)
            elif os.path.isdir(path):
                shutil.rmtree(path)
        except OSError as e:
            print(f"  Error during cleanup of {path}: {e}")

# --- PART 1: SCRAPING TENDER INFORMATION ---
print("\n--- PART 1: STARTING BROWSER AND SCRAPING ---")
driver = None
try:
    print("Initializing WebDriver...")
    driver = webdriver.Chrome(options=chrome_options)
    print("WebDriver started successfully.")
    
    driver.get("https://www.offresonline.com/")
    time.sleep(2)
    driver.find_element(By.CSS_SELECTOR, "#main-nav > ul > li:nth-child(2)").click()
    time.sleep(1)
    driver.find_element(By.CSS_SELECTOR, "#Login").send_keys(OFFRES_USERNAME)
    driver.find_element(By.CSS_SELECTOR, "#pwd").send_keys(OFFRES_PASSWORD)
    driver.find_element(By.CSS_SELECTOR, "#buuuttt").click()
    print("Login successful.")
    time.sleep(2)
    
    driver.find_element(By.CSS_SELECTOR, "#ctl00_Linkf30").click()
    print("Navigating to 'AO de Jour' (Daily Tenders)...")
    
    table_body = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//*[@id="tableao"]/tbody')))
    rows = table_body.find_elements(By.TAG_NAME, "tr")
    print(f"Found {len(rows)} rows in the tender table.")

    all_data = []
    for row in rows:
        td_list = row.find_elements(By.TAG_NAME, "td")
        if len(td_list) < 6: continue
        
        soup = bs(td_list[2].get_attribute('innerHTML'), 'html.parser')
        strong_tags = soup.find_all('strong')
        organisme = strong_tags[0].get_text(strip=True) if strong_tags else ''
        objet = strong_tags[1].get_text(strip=True) if len(strong_tags) > 1 else ''
        
        value = ""
        try:
            input_el = row.find_element(By.XPATH, ".//input[@title='Marquer comme Lu ?' or @title='Déja vu']")
            value = input_el.get_attribute("value")
        except NoSuchElementException: continue
        
        if objet and value:
            all_data.append({'Objet': objet, 'Value': value, 'Organisme': organisme})

    if not all_data:
        raise Exception("Scraping completed, but no tender data was collected.")

    df = pd.DataFrame(all_data)
    print(f"Created DataFrame with {len(df)} initial tenders.")

    excluded_words = ["construction", "installation", "recrutement", "travaux", "fourniture", "achat", "equipement", "maintenance", "works", "goods", "supply", "acquisition", "recruitment", "nettoyage", "gardiennage"]
    df = df[~df['Objet'].str.lower().str.contains('|'.join(excluded_words), na=False)].reset_index(drop=True)
    print(f"Filtered to {len(df)} relevant tenders.")

except Exception as e:
    print(f"FATAL ERROR during scraping: {e}")
    if driver: driver.quit()
    raise

# --- PART 2: SEQUENTIAL PROCESSING (DOWNLOAD, EXTRACT, SEND) ---
print("\n--- PART 2: STARTING SEQUENTIAL TENDER PROCESSING ---")
client = OpenAI(api_key=OPENAI_API_KEY)

if df.empty:
    print("No relevant tenders to process. Exiting.")
else:
    for index, row in df.iterrows():
        number, objet, organisme = row['Value'], row['Objet'], row['Organisme']
        print(f"\n--- Processing Tender {index + 1}/{len(df)} | Value: {number} ---")
        
        paths_to_clean = []
        try:
            # Step 1: Download the file(s)
            files_before = set(os.listdir(download_folder))
            url = f'https://www.offresonline.com/Admin/telechargercps.aspx?http=N&i={number}&type=1&encour=1&p=p'
            driver.get(url)
            time.sleep(1)

            png = driver.find_element(By.TAG_NAME, "body").screenshot_as_png
            b64 = base64.b64encode(png).decode("utf-8")
            
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[{"role": "user", "content": [{"type": "text", "text": "Transcribe the characters in the image. Respond with ONLY the characters."}, {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}}]}],
                max_tokens=10
            )
            captcha_code = response.choices[0].message.content.strip()
            print(f"  OCR Result: '{captcha_code}'")

            driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtimgcode").send_keys(captcha_code)
            driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_LinkButton1").click()
            print("  Download initiated. Waiting for completion...")
            
            if not wait_for_download_complete(download_folder):
                continue
            
            # Step 2: Identify and Extract Text from new file(s)
            files_after = set(os.listdir(download_folder))
            new_files = files_after - files_before
            if not new_files:
                print("  No new file detected after download. Skipping.")
                continue

            merged_text = ""
            for new_file_name in new_files:
                file_path = os.path.join(download_folder, new_file_name)
                paths_to_clean.append(file_path)
                
                if new_file_name.lower().endswith(".zip"):
                    extract_dir = os.path.join(download_folder, os.path.splitext(new_file_name)[0])
                    paths_to_clean.append(extract_dir)
                    with zipfile.ZipFile(file_path, 'r') as zf:
                        zf.extractall(extract_dir)
                    print(f"  Unzipped '{new_file_name}'.")
                    
                    for root, _, files in os.walk(extract_dir):
                        for f in files:
                            if 'cps' in f.lower(): continue
                            text = process_file_for_text(os.path.join(root, f))
                            if text: merged_text += f"\n\n--- Content from: {f} ---\n{text}"
                else: # Not a zip file
                     if 'cps' not in new_file_name.lower():
                        text = process_file_for_text(file_path)
                        if text: merged_text += f"\n\n--- Content from: {new_file_name} ---\n{text}"
            
            # Step 3: Send Data to n8n Webhook
            if merged_text.strip():
                print(f"  Extracted {len(merged_text)} characters. Sending to webhook...")
                payload = {
                    'Objet': objet,
                    'Value': number,
                    'Organisme': organisme,
                    'merged_text': merged_text.strip()
                }
                # n8n expects a list of items, so we send a list containing our single payload
                response = requests.post(N8N_WEBHOOK_URL, json=[payload], timeout=30)
                response.raise_for_status()
                print(f"  ✅ SUCCESS: Tender {number} sent to n8n.")
            else:
                print(f"  ⚠️ No text extracted for tender {number}. Nothing to send.")

        except Exception as e:
            print(f"  ❌ FAILED to process tender {number}: {e}")
        
        finally:
            # Step 4: Clean up and Wait
            print(f"  Cleaning up files for tender {number}...")
            cleanup_files(paths_to_clean)
            print("  Waiting 10 seconds before next tender...")
            time.sleep(10)

if driver:
    driver.quit()
print("\nWebDriver closed.")
print("\n--- SCRIPT FINISHED ---")
