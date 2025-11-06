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
from bs4 import BeautifulSoup as bs
import signal  # ✅ for timeout handling

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

from openai import OpenAI

print("--- SCRIPT STARTED ---")

# --- Credentials and Configuration ---
OFFRES_USERNAME = "TARGETUP"
OFFRES_PASSWORD = "TARGETUP2024"
OPENAI_API_KEY = "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"
N8N_WEBHOOK_URL = "https://targetup.app.n8n.cloud/webhook/dc4cf7c8-b44e-4404-830d-ef7cf3e7b6ca"

print("Credentials and configuration loaded.")

# --- Timeout helper ---
class TimeoutException(Exception):
    pass

def timeout(seconds=300):
    """Decorator to stop function after `seconds`."""
    def decorator(func):
        def wrapper(*args, **kwargs):
            def handler(signum, frame):
                raise TimeoutException(f"Timed out after {seconds} seconds")
            signal.signal(signal.SIGALRM, handler)
            signal.alarm(seconds)
            try:
                return func(*args, **kwargs)
            finally:
                signal.alarm(0)
        return wrapper
    return decorator

# --- Setup download folder ---
workspace_path = os.getcwd()
download_folder = os.path.join(workspace_path, "downloads")
os.makedirs(download_folder, exist_ok=True)
print(f"Download folder created at: {download_folder}")

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

# --- PART 1: SCRAPING ---
print("\n--- PART 1: STARTING BROWSER AND SCRAPING ---")
try:
    print("Initializing WebDriver...")
    driver = webdriver.Chrome(options=chrome_options)
    print("WebDriver started successfully in headless mode.")

    driver.get("https://www.offresonline.com/")
    time.sleep(2)

    print("Clicking 'Se connecter'...")
    driver.find_element(By.CSS_SELECTOR, "#main-nav > ul > li:nth-child(2)").click()
    time.sleep(1)

    print("Entering login credentials...")
    driver.find_element(By.CSS_SELECTOR, "#Login").send_keys(OFFRES_USERNAME)
    time.sleep(1)
    driver.find_element(By.CSS_SELECTOR, "#pwd").send_keys(OFFRES_PASSWORD)
    time.sleep(1)
    driver.find_element(By.CSS_SELECTOR, "#buuuttt").click()
    print("Login successful.")
    time.sleep(2)

    print("Navigating to 'AO de Jour' (Daily Tenders)...")
    driver.find_element(By.CSS_SELECTOR, "#ctl00_Linkf30").click()

    print("Waiting for the tender table to load...")
    table_body = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="tableao"]/tbody'))
    )
    rows = table_body.find_elements(By.TAG_NAME, "tr")
    print(f"Found {len(rows)} rows in the tender table.")

    all_data = []
    for i, row in enumerate(rows):
        try:
            td_list = row.find_elements(By.TAG_NAME, "td")
            if len(td_list) < 6:
                print(f"  Row {i+1}: Skipping malformed row.")
                continue

            organisme_objet_html = td_list[2].get_attribute('innerHTML')
            soup = bs(organisme_objet_html, 'html.parser')
            strong_tags = soup.find_all('strong')
            objet = strong_tags[1].get_text(strip=True) if len(strong_tags) > 1 else ''
            value = ""
            try:
                input_element = row.find_element(By.XPATH, ".//input[@title='Marquer comme Lu ?' or @title='Déja vu']")
                value = input_element.get_attribute("value")
            except NoSuchElementException:
                value = ""

            all_data.append({'Objet': objet, 'Value': value})
        except Exception as e:
            print(f"  Row {i+1}: Skipped due to error: {e}")
            continue

    if not all_data:
        print("WARNING: No data scraped. Exiting.")
        driver.quit()
        exit()

    df = pd.DataFrame(all_data)
    print(f"Created DataFrame with {len(df)} rows.")

    excluded_words = [
        "construction", "installation", "recrutement", "travaux", "fourniture", "achat",
        "equipement", "maintenance", "works", "goods", "supply", "acquisition",
        "recruitment", "nettoyage", "gardiennage"
    ]
    df = df[~df['Objet'].str.lower().str.contains('|'.join(excluded_words), na=False)]
    df = df[df['Value'] != ''].reset_index(drop=True)
    print(f"Filtered to {len(df)} relevant rows.")

except Exception as e:
    print(f"FATAL ERROR during scraping: {e}")
    driver.save_screenshot("error_page_scraping.png")
    with open("error_page_scraping.html", "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    driver.quit()
    raise

# --- PART 2: DOWNLOAD FILES ---
print("\n--- PART 2: DOWNLOADING FILES ---")
client = OpenAI(api_key=OPENAI_API_KEY)
if df.empty:
    print("No tenders to download after filtering.")
    driver.quit()
    exit()

for index, row in df.iterrows():
    number = row['Value']
    objet = row['Objet']
    print(f"\nProcessing download {index + 1}/{len(df)} | Value: {number} | Objet: {objet[:50]}...")
    try:
        url = f'https://www.offresonline.com/Admin/telechargercps.aspx?http=N&i={number}&type=1&encour=1&p=p'
        driver.get(url)
        time.sleep(1)
        print("  Taking screenshot for CAPTCHA...")
        png = driver.find_element(By.TAG_NAME, "body").screenshot_as_png
        b64 = base64.b64encode(png).decode("utf-8")

        captcha_code = ""
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "text", "text": "You are an OCR system. Respond only with the text you see."},
                        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}}
                    ]
                }],
                max_tokens=10
            )
            captcha_code = response.choices[0].message.content.strip()
            print(f"  OCR result for {number}: '{captcha_code}'")
        except Exception as ocr_e:
            print(f"  ⚠️ OCR failed: {ocr_e}")
            captcha_code = ""

        captcha_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_txtimgcode"))
        )
        captcha_input.send_keys(captcha_code)
        driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_LinkButton1").click()
        print(f"  Download triggered for {number}. Waiting for file...")
        time.sleep(15)

    except Exception as e:
        print(f"  ERROR downloading {number}: {e}")
        continue

driver.quit()
print("WebDriver closed. Downloads complete.")

# --- PART 3: EXTRACTING FILES ---
print("\n--- PART 3: EXTRACTING TEXT ---")

@timeout(120)
def extract_text_from_pdf(file_path):
    text = ""
    try:
        doc = fitz.open(file_path)
        for page in doc:
            text += page.get_text("text") + "\n"
        doc.close()
        if len(text.strip()) < 100:
            print("    -> Short PDF text. Performing OCR...")
            pages = convert_from_path(file_path)
            ocr_text = ""
            for p_img in pages:
                ocr_text += pytesseract.image_to_string(p_img, lang="fra+ara") + "\n"
            return ocr_text.strip()
        return text.strip()
    except Exception as e:
        print(f"    -> ERROR reading PDF {os.path.basename(file_path)}: {e}")
        return ""

@timeout(120)
def extract_from_zip(file_path, tenders_dir):
    extract_to = os.path.join(tenders_dir, os.path.splitext(os.path.basename(file_path))[0])
    os.makedirs(extract_to, exist_ok=True)
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)
    print(f"  Unzipped '{os.path.basename(file_path)}'.")
    os.remove(file_path)

def extract_text_from_docx(file_path):
    try:
        doc = docx.Document(file_path)
        return "\n".join([p.text for p in doc.paragraphs])
    except Exception as e:
        print(f"    -> ERROR reading DOCX {os.path.basename(file_path)}: {e}")
        return ""

# Unzip all ZIPs
print("\nStep 1: Unzipping all downloaded ZIP files...")
for f in os.listdir(download_folder):
    if f.lower().endswith(".zip"):
        try:
            extract_from_zip(os.path.join(download_folder, f), download_folder)
        except TimeoutException as te:
            print(f"⚠️ Timeout unzipping {f}: {te}. Skipping.")
        except Exception as e:
            print(f"⚠️ Failed to unzip {f}: {e}")

# Extract text
print("\nStep 2: Extracting text from files...")
tender_results = []
items_in_downloads = os.listdir(download_folder)
for item_name in items_in_downloads:
    item_path = os.path.join(download_folder, item_name)
    merged_text = ""
    print(f"\nProcessing item: '{item_name}'")
    try:
        if os.path.isdir(item_path):
            for root, _, files in os.walk(item_path):
                for f in files:
                    if 'cps' in f.lower():
                        print(f"  -> Skipping 'cps' file: {f}")
                        continue
                    file_path = os.path.join(root, f)
                    ext = os.path.splitext(f)[1].lower()
                    text = ""
                    print(f"  -> Reading file: {f}")
                    try:
                        if ext == ".pdf":
                            text = extract_text_from_pdf(file_path)
                        elif ext == ".docx":
                            text = extract_text_from_docx(file_path)
                    except TimeoutException as te:
                        print(f"  ⚠️ Timeout on {f}: {te}. Skipping.")
                        continue
                    if text.strip():
                        merged_text += f"\n\n--- Content from: {f} ---\n{text}"
        if merged_text.strip():
            tender_results.append({"tender_folder": item_name, "merged_text": merged_text.strip()})
            print(f"  Finished '{item_name}', extracted {len(merged_text)} characters.")
    except TimeoutException as te:
        print(f"⚠️ Timeout processing folder {item_name}: {te}. Skipping this tender.")
        continue
    except Exception as e:
        print(f"⚠️ Error processing {item_name}: {e}")

# --- PART 4: SEND TO WEBHOOK ---
print("\n--- PART 4: FINALIZING ---")
if tender_results:
    final_df = pd.DataFrame(tender_results)
    print(f"Created DataFrame with {len(final_df)} tenders.")
    final_df.to_csv("tender_results.csv", index=False)
    print("Saved to tender_results.csv.")
    try:
        json_data = final_df.to_dict(orient="records")
        response = requests.post(N8N_WEBHOOK_URL, json=json_data)
        if response.status_code == 200:
            print("✅ Data sent successfully to n8n webhook!")
        else:
            print(f"❌ Webhook error: {response.status_code} | {response.text}")
    except Exception as e:
        print(f"⚠️ Failed to send data: {e}")
else:
    print("No text extracted. Nothing to send.")

print("\n--- SCRIPT FINISHED ---")
