"""
Calcutta High Court - Jalpaiguri Circuit Bench Display Board Scraper
URL: https://display.calcuttahighcourt.gov.in/jalpaiguri.php
Features: Auto CAPTCHA detection + Manual entry fallback + API Integration + Excel Backup
"""

import time
import os
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import platform
import requests
import json
import cv2
import numpy as np
import pytesseract
from PIL import Image
import io
import base64

# ==================== CONFIGURATION ====================
URL = "https://display.calcuttahighcourt.gov.in/portblair.php"
SCRAPE_INTERVAL = 30  # seconds
BASE_FOLDER = r"D:\CourtDisplayBoardScraper\displayboardexcel\calcutta_hc_excel\portblair_hc_detailed_excel"
BACKUP_CYCLE_INTERVAL = 60  # Create backup after every 60 cycles
BENCH_NAME = "Port Blair"

# API Configuration
API_URL = "https://api.courtlivestream.com/api/display-boards/create"
API_TIMEOUT = 10  # seconds
ENABLE_API_POSTING = True
ENABLE_EXCEL_SAVING = True

# CAPTCHA Configuration
MANUAL_CAPTCHA_TIMEOUT = 30  # seconds to wait for manual entry
AUTO_CAPTCHA_ENABLED = True  # Try auto-detection first

# Tesseract path (update this to your installation path)
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# ==================== HELPER FUNCTIONS ====================

def extract_case_number_numeric(case_full):
    """Extract numeric part from case number like MAT/67/2026 -> 67"""
    try:
        if not case_full or not case_full.strip():
            return ""
        
        # Pattern: Find number between slashes or before slash
        # MAT/67/2026 -> 67
        match = re.search(r'/(\d+)/', case_full)
        if match:
            return match.group(1).strip()
        
        # Fallback: WPA/444/2026 -> 444
        match = re.search(r'/(\d+)', case_full)
        if match:
            return match.group(1).strip()
        
        return ""
    except:
        return ""


def extract_serial_number_range(serial_str):
    """
    Extract serial number range from strings like:
    - 'AD 7' -> [7]
    - 'AD 27-31' -> [27, 28, 29, 30, 31]
    - 'OD 10-11' -> [10, 11]
    """
    try:
        if not serial_str or not serial_str.strip():
            return []
        
        # Check for range pattern: AD 27-31
        range_match = re.search(r'(\d+)-(\d+)', serial_str)
        if range_match:
            start = int(range_match.group(1))
            end = int(range_match.group(2))
            return list(range(start, end + 1))
        
        # Single number: AD 7
        single_match = re.search(r'\b(\d+)\b', serial_str)
        if single_match:
            return [int(single_match.group(1))]
        
        return []
    except:
        return []


# ==================== CAPTCHA FUNCTIONS ====================

def download_captcha_image(driver):
    """Download CAPTCHA image from the page"""
    try:
        captcha_img_element = driver.find_element(By.CSS_SELECTOR, "#captcha_image img")
        captcha_src = captcha_img_element.get_attribute("src")
        
        # Get image as bytes
        img_base64 = driver.execute_script("""
            var img = arguments[0];
            var canvas = document.createElement('canvas');
            canvas.width = img.width;
            canvas.height = img.height;
            var ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0);
            return canvas.toDataURL('image/png').substring(22);
        """, captcha_img_element)
        
        img_bytes = io.BytesIO(base64.b64decode(img_base64))
        return Image.open(img_bytes)
        
    except Exception as e:
        print(f"   ✗ Error downloading CAPTCHA: {str(e)}")
        return None


def preprocess_captcha_image(image):
    """Preprocess CAPTCHA image for better OCR"""
    try:
        # Convert PIL to OpenCV format
        img_array = np.array(image)
        
        # Convert to grayscale
        gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
        
        # Apply thresholding
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
        
        # Denoise
        denoised = cv2.fastNlMeansDenoising(thresh)
        
        # Resize for better OCR
        resized = cv2.resize(denoised, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
        
        return resized
        
    except Exception as e:
        print(f"   ✗ Error preprocessing image: {str(e)}")
        return None


def detect_captcha_text(image):
    """Use OCR to detect CAPTCHA text"""
    try:
        # Preprocess image
        processed = preprocess_captcha_image(image)
        if processed is None:
            return None
        
        # Configure tesseract for alphanumeric only
        custom_config = r'--oem 3 --psm 7 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'
        
        # Extract text
        text = pytesseract.image_to_string(processed, config=custom_config)
        text = text.strip().replace(" ", "")
        
        # Validate: should be 5-6 alphanumeric characters
        if text and len(text) >= 5 and text.isalnum():
            return text
        
        return None
        
    except Exception as e:
        print(f"   ✗ OCR Error: {str(e)}")
        return None


def try_auto_captcha(driver):
    """Attempt automatic CAPTCHA detection and validation"""
    if not AUTO_CAPTCHA_ENABLED:
        return False
    
    try:
        print("\n" + "="*100)
        print("ATTEMPTING AUTOMATIC CAPTCHA DETECTION")
        print("="*100)
        
        # Download CAPTCHA image
        print("   → Downloading CAPTCHA image...")
        captcha_image = download_captcha_image(driver)
        
        if captcha_image is None:
            print("   ✗ Failed to download CAPTCHA image")
            return False
        
        print("   ✓ CAPTCHA image downloaded")
        
        # Detect CAPTCHA text
        print("   → Running OCR on CAPTCHA...")
        captcha_text = detect_captcha_text(captcha_image)
        
        if captcha_text is None:
            print("   ✗ OCR failed to detect CAPTCHA text")
            return False
        
        print(f"   ✓ OCR Detected: '{captcha_text}'")
        
        # Enter CAPTCHA
        print("   → Entering CAPTCHA...")
        security_code_input = driver.find_element(By.ID, "security_code")
        security_code_input.clear()
        security_code_input.send_keys(captcha_text)
        
        # Click validate button
        print("   → Clicking Validate CAPTCHA button...")
        validate_button = driver.find_element(By.ID, "validate_captcha")
        validate_button.click()
        
        # Wait for validation
        time.sleep(3)
        
        # Check if CAPTCHA div is hidden (success)
        captcha_div = driver.find_element(By.ID, "captcha_div")
        if captcha_div.get_attribute("style") == "display: none;":
            print(f"\n{'='*100}")
            print(f"✓✓✓ AUTOMATIC CAPTCHA VALIDATION SUCCESSFUL ✓✓✓")
            print(f"   Detected Text: {captcha_text}")
            print(f"{'='*100}\n")
            return True
        else:
            print("   ✗ CAPTCHA validation failed - incorrect text detected")
            return False
            
    except Exception as e:
        print(f"   ✗ Auto CAPTCHA error: {str(e)}")
        return False


def handle_manual_captcha(driver):
    """Wait for manual CAPTCHA entry"""
    print("\n" + "="*100)
    print("WAITING FOR MANUAL CAPTCHA ENTRY")
    print("="*100)
    print(f"   ⚠ Please enter CAPTCHA manually within {MANUAL_CAPTCHA_TIMEOUT} seconds")
    print(f"   ⚠ Click 'Validate CAPTCHA' button after entering")
    print("="*100)
    
    start_time = time.time()
    
    while time.time() - start_time < MANUAL_CAPTCHA_TIMEOUT:
        try:
            # Check if CAPTCHA div is hidden
            captcha_div = driver.find_element(By.ID, "captcha_div")
            if captcha_div.get_attribute("style") == "display: none;":
                print(f"\n{'='*100}")
                print(f"✓✓✓ MANUAL CAPTCHA VALIDATION SUCCESSFUL ✓✓✓")
                print(f"{'='*100}\n")
                return True
            
            time.sleep(1)
            
        except Exception as e:
            continue
    
    print(f"\n{'='*100}")
    print(f"✗ MANUAL CAPTCHA TIMEOUT - Please try again")
    print(f"{'='*100}\n")
    return False


def validate_captcha(driver):
    """Main CAPTCHA validation function - tries auto then manual"""
    
    # Try automatic CAPTCHA detection first
    if try_auto_captcha(driver):
        return True
    
    # If auto fails, wait for manual entry
    print("\n   ⚠ Automatic CAPTCHA detection failed")
    print("   ⚠ Switching to manual CAPTCHA entry mode...")
    
    return handle_manual_captcha(driver)


# ==================== API FUNCTIONS ====================

def post_court_data_to_api(court_data):
    """Post a single court record to the API"""
    try:
        datetime_str = court_data.get("DateTime", "")
        
        if datetime_str:
            try:
                dt_obj = datetime.strptime(datetime_str, "%Y-%m-%d %H:%M:%S")
                date_str = dt_obj.strftime("%Y-%m-%d")
                time_str = dt_obj.strftime("%I:%M %p")
            except ValueError:
                date_str = datetime.now().strftime("%Y-%m-%d")
                time_str = datetime.now().strftime("%I:%M %p")
        else:
            date_str = datetime.now().strftime("%Y-%m-%d")
            time_str = datetime.now().strftime("%I:%M %p")
        
        court_no = court_data.get("Court", "")
        case_number = court_data.get("Case Number", "")
        serial_no = court_data.get("Serial No(s)", "")
        
        # Convert serial number to integer
        try:
            serial_number = int(serial_no) if serial_no else 0
        except (ValueError, TypeError):
            serial_number = 0
        
        payload = {
            "benchName": BENCH_NAME,
            "courtHallNumber": court_no,
            "caseNumber": case_number,
            "serialNumber": serial_number,
            "date": date_str,
            "time": time_str,
            "stage": court_data.get("Judge(s) Coram", ""),
            "listNumber": 0
        }
        
        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        
        response = requests.post(
            API_URL,
            json=payload,
            headers=headers,
            timeout=API_TIMEOUT
        )
        
        if response.status_code in [200, 201]:
            return True, response.json()
        else:
            return False, f"API Error {response.status_code}: {response.text}"
    
    except requests.exceptions.Timeout:
        return False, "Request timeout"
    except requests.exceptions.ConnectionError:
        return False, "Connection error"
    except Exception as e:
        return False, f"Error: {str(e)}"


def post_all_courts_to_api(courts_data_list):
    """Post all court records to the API"""
    if not ENABLE_API_POSTING:
        print("\n   ⚠ API posting is DISABLED in configuration")
        return {"total": 0, "successful": 0, "failed": 0, "errors": []}
    
    total_courts = len(courts_data_list)
    successful_posts = 0
    failed_posts = 0
    errors = []
    
    print(f"\n{'='*100}")
    print(f"POSTING {total_courts} COURT RECORDS TO API")
    print(f"API URL: {API_URL}")
    print(f"MAPPING: Court->courtHallNumber | Case Number->caseNumber | Serial No->serialNumber")
    print(f"{'='*100}\n")
    
    for idx, court_data in enumerate(courts_data_list, 1):
        court_no = court_data.get("Court", "N/A")
        case_num = court_data.get("Case Number", "N/A")
        serial_no = court_data.get("Serial No(s)", "N/A")
        
        print(f"   [{idx}/{total_courts}] Court={court_no} | Case={case_num} | Serial={serial_no}", end=" ")
        
        success, response = post_court_data_to_api(court_data)
        
        if success:
            successful_posts += 1
            print("✓")
        else:
            failed_posts += 1
            print(f"✗ ({response})")
            errors.append({"court": court_no, "case": case_num, "error": response})
    
    print(f"\n{'='*100}")
    print(f"API POSTING SUMMARY")
    print(f"{'='*100}")
    print(f"   Total: {total_courts}")
    print(f"   ✓ Successful: {successful_posts}")
    print(f"   ✗ Failed: {failed_posts}")
    print(f"   Success rate: {(successful_posts/total_courts*100):.1f}%")
    
    if errors and len(errors) > 0:
        print(f"\n   FAILED RECORDS (showing first 5):")
        for err in errors[:5]:
            print(f"      - Court {err['court']}: {err['error']}")
        if len(errors) > 5:
            print(f"      ... and {len(errors)-5} more errors")
    
    print(f"{'='*100}\n")
    
    return {
        "total": total_courts,
        "successful": successful_posts,
        "failed": failed_posts,
        "errors": errors
    }


# ==================== SETUP FUNCTIONS ====================

def setup_driver():
    """Initialize Chrome driver with VISIBLE browser"""
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager
    
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.implicitly_wait(10)
    return driver


def create_folder():
    """Create date-based folder structure"""
    if not ENABLE_EXCEL_SAVING:
        return None
        
    current_date = datetime.now().strftime("%Y_%m_%d")
    date_folder = os.path.join(BASE_FOLDER, f"jalpaiguri_hc_detailed_{current_date}")
    
    if not os.path.exists(date_folder):
        os.makedirs(date_folder)
        print(f"✓ Created folder: {date_folder}")
    
    return date_folder


def get_date_folder():
    """Get today's date-based folder path"""
    current_date = datetime.now().strftime("%Y_%m_%d")
    date_folder = os.path.join(BASE_FOLDER, f"jalpaiguri_hc_detailed_{current_date}")
    return date_folder


def get_excel_path(folder):
    """Get full path for today's main Excel file"""
    if not folder:
        return None
    current_date = datetime.now().strftime("%Y_%m_%d")
    filename = f"jalpaiguri_hc_detailed_{current_date}.xlsx"
    excel_path = os.path.join(folder, filename)
    return excel_path


def get_timestamped_backup_path(folder):
    """Get full path for timestamped backup Excel file"""
    if not folder:
        return None
    current_timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M")
    filename = f"jalpaiguri_hc_detailed_bk_{current_timestamp}.xlsx"
    backup_path = os.path.join(folder, filename)
    return backup_path


def create_backup_from_main_excel(main_excel_path, folder):
    """Create a timestamped backup file"""
    if not ENABLE_EXCEL_SAVING or not main_excel_path or not folder:
        return False
        
    try:
        if not os.path.exists(main_excel_path):
            print("   ⚠ Main Excel file does not exist yet. Cannot create backup.")
            return False
        
        main_df = pd.read_excel(main_excel_path, engine='openpyxl')
        
        if main_df.empty:
            print("   ⚠ Main Excel file is empty. No backup created.")
            return False
        
        backup_path = get_timestamped_backup_path(folder)
        main_df.to_excel(backup_path, index=False, sheet_name='Sheet1', engine='openpyxl')
        
        print(f"\n{'='*100}")
        print(f"✓✓✓ TIMESTAMPED BACKUP CREATED ✓✓✓")
        print(f"   Backup file: {os.path.basename(backup_path)}")
        print(f"   Total rows backed up: {len(main_df)}")
        print(f"   Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"{'='*100}\n")
        
        return True
        
    except Exception as e:
        print(f"   ✗ Error creating backup: {str(e)}")
        return False


def open_excel_file(file_path):
    """Open Excel file automatically"""
    try:
        if platform.system() == 'Windows':
            os.startfile(file_path)
            print(f"   ✓ Excel file opened: {file_path}")
    except Exception as e:
        print(f"   ⚠ Could not auto-open Excel: {str(e)}")


# ==================== SCRAPING FUNCTIONS ====================

def extract_case_numbers_from_eye_button(driver, row_element):
    """Click eye button and extract case numbers from popup"""
    try:
        # Find eye button in this row
        eye_button = row_element.find_element(By.CSS_SELECTOR, "span[onclick*='viewCases'] i.fa-eye")
        
        # Get the onclick attribute to extract case numbers
        parent_span = eye_button.find_element(By.XPATH, "..")
        onclick_attr = parent_span.get_attribute("onclick")
        
        # Extract case numbers from onclick attribute
        # Format: viewCases('AD 7','MAT/67/2026')
        match = re.search(r"viewCases\('([^']+)','([^']+)'\)", onclick_attr)
        
        if match:
            serial_display = match.group(1).strip()  # e.g., "AD 7"
            case_numbers_str = match.group(2).strip()  # e.g., "MAT/67/2026" or "WPA/444/2026, WPA/445/2026"
            
            # Split multiple case numbers and clean whitespace
            case_numbers = [cn.strip() for cn in case_numbers_str.split(",")]
            
            return serial_display, case_numbers
        
        return None, []
        
    except Exception as e:
        return None, []


def scrape_display_board(driver):
    """Scrape Jalpaiguri HC display board table"""
    try:
        print("\n" + "="*100)
        print("SCRAPING DISPLAY BOARD")
        print("="*100)
        
        scrape_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Wait for table to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "display-board-table"))
        )
        
        time.sleep(2)  # Extra wait for data to populate
        
        all_courts_data = []
        
        # Find all table rows in tbody
        table = driver.find_element(By.ID, "display-board-table")
        tbody = table.find_element(By.TAG_NAME, "tbody")
        rows = tbody.find_elements(By.TAG_NAME, "tr")
        
        print(f"   → Found {len(rows)} court rows in table\n")
        
        for idx, row in enumerate(rows, 1):
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                
                if len(cells) < 3:
                    continue
                
                # Extract Court Number
                court_cell = cells[0]
                court_no_span = court_cell.find_elements(By.TAG_NAME, "span")
                if len(court_no_span) > 0:
                    court_no = court_cell.text.replace("ℹ", "").strip()
                else:
                    court_no = court_cell.text.strip()
                
                # Extract Judge(s) Coram
                judges = cells[1].text.strip()
                
                # Extract Serial No and Case Numbers from eye button
                serial_display_full = cells[2].text.strip()
                
                # Get serial number range (e.g., "AD 27-31" -> [27, 28, 29, 30, 31])
                serial_numbers = extract_serial_number_range(serial_display_full)
                
                # Extract case numbers from eye button
                serial_display, case_numbers = extract_case_numbers_from_eye_button(driver, row)
                
                if case_numbers and serial_numbers:
                    # Map each serial number to corresponding case number
                    # If counts don't match, we'll handle gracefully
                    num_records = max(len(serial_numbers), len(case_numbers))
                    
                    for i in range(num_records):
                        # Get serial number (use last if we run out)
                        if i < len(serial_numbers):
                            serial_no = serial_numbers[i]
                        else:
                            serial_no = serial_numbers[-1] if serial_numbers else ""
                        
                        # Get case number (use last if we run out)
                        if i < len(case_numbers):
                            case_full = case_numbers[i]
                        else:
                            case_full = case_numbers[-1] if case_numbers else ""
                        
                        case_number_numeric = extract_case_number_numeric(case_full)
                        
                        court_data = {
                            "Bench Name": BENCH_NAME,
                            "Court": court_no,
                            "Judge(s) Coram": judges,
                            "Serial No(s)": str(serial_no),
                            "Case Number (Full)": case_full,
                            "Case Number": case_number_numeric,
                            "DateTime": scrape_time
                        }
                        
                        all_courts_data.append(court_data)
                        print(f"      ✓ Court {court_no}: Serial {serial_no} | Case {case_number_numeric} ({case_full})")
                
                elif serial_numbers and not case_numbers:
                    # Serial numbers but no cases - create records without case data
                    for serial_no in serial_numbers:
                        court_data = {
                            "Bench Name": BENCH_NAME,
                            "Court": court_no,
                            "Judge(s) Coram": judges,
                            "Serial No(s)": str(serial_no),
                            "Case Number (Full)": "",
                            "Case Number": "",
                            "DateTime": scrape_time
                        }
                        
                        all_courts_data.append(court_data)
                        print(f"      ⚠ Court {court_no}: Serial {serial_no} | No case numbers")
                
                else:
                    # No serial numbers found - fallback
                    print(f"      ⚠ Court {court_no}: Could not extract serial numbers")
                
            except Exception as e:
                print(f"      ✗ Error processing row {idx}: {str(e)}")
                continue
        
        print(f"\n{'='*100}")
        print(f"EXTRACTION SUMMARY:")
        print(f"{'='*100}")
        print(f"   ✓ Total records extracted: {len(all_courts_data)}")
        print(f"   ✓ Timestamp: {scrape_time}")
        
        # Show example of range mapping
        if all_courts_data:
            court_grouped = {}
            for d in all_courts_data:
                court_key = d['Court']
                if court_key not in court_grouped:
                    court_grouped[court_key] = []
                court_grouped[court_key].append(d)
            
            # Check if any court has multiple records (range mapping)
            has_ranges = any(len(records) > 1 for records in court_grouped.values())
            
            if has_ranges:
                print(f"\n   Sample Range Mappings:")
                shown = 0
                for court_key, records in court_grouped.items():
                    if len(records) > 1 and shown < 3:  # Show first 3 courts with ranges
                        print(f"      Court {court_key}:")
                        for rec in records[:5]:  # Show first 5 records
                            print(f"         Serial {rec['Serial No(s)']} → Case {rec['Case Number']} ({rec['Case Number (Full)']})")
                        if len(records) > 5:
                            print(f"         ... and {len(records)-5} more")
                        shown += 1
        
        print(f"{'='*100}\n")
        
        return all_courts_data
    
    except Exception as e:
        print(f"\n   ✗ ERROR during scraping: {str(e)}")
        import traceback
        traceback.print_exc()
        return []


# ==================== EXCEL SAVE FUNCTIONS ====================

def save_to_excel(data, file_path, open_file=False):
    """Save scraped data to Excel file"""
    if not ENABLE_EXCEL_SAVING or not file_path:
        print("\n   ⚠ Excel saving is DISABLED")
        return False
        
    try:
        if not data:
            print("   ⚠ No data to save")
            return False
        
        df = pd.DataFrame(data)
        df = df[["Bench Name", "Court", "Judge(s) Coram", "Serial No(s)", "Case Number (Full)", "Case Number", "DateTime"]]
        
        if os.path.exists(file_path):
            existing_df = pd.read_excel(file_path, engine='openpyxl')
            combined_df = pd.concat([existing_df, df], ignore_index=True)
            combined_df.to_excel(file_path, index=False, sheet_name='Sheet1', engine='openpyxl')
            
            print(f"\n   ✓ Data appended to Excel")
            print(f"   ✓ Added {len(df)} records (Total: {len(combined_df)} rows)")
            print(f"   ✓ File: {os.path.basename(file_path)}")
        else:
            df.to_excel(file_path, index=False, sheet_name='Sheet1', engine='openpyxl')
            print(f"\n   ✓ New Excel file created")
            print(f"   ✓ Initial data: {len(df)} records")
            print(f"   ✓ File: {os.path.basename(file_path)}")
        
        if open_file:
            open_excel_file(file_path)
        
        return True
        
    except Exception as e:
        print(f"   ✗ Error saving to Excel: {str(e)}")
        return False


# ==================== MAIN EXECUTION ====================

def main():
    """Main execution"""
    print("=" * 100)
    print(" " * 10 + "CALCUTTA HIGH COURT - PORT BLAIR CIRCUIT BENCH DISPLAY BOARD SCRAPER")
    print(" " * 15 + "WITH AUTO CAPTCHA + EXCEL BACKUP + API INTEGRATION")
    print("=" * 100)
    print(f"URL: {URL}")
    print(f"Bench Name: {BENCH_NAME}")
    print(f"Auto CAPTCHA: {'ENABLED' if AUTO_CAPTCHA_ENABLED else 'DISABLED'}")
    print(f"Manual CAPTCHA Timeout: {MANUAL_CAPTCHA_TIMEOUT} seconds")
    print(f"Scrape Interval: {SCRAPE_INTERVAL} seconds")
    print(f"Excel Saving: {'ENABLED' if ENABLE_EXCEL_SAVING else 'DISABLED'}")
    if ENABLE_EXCEL_SAVING:
        print(f"Base Location: {BASE_FOLDER}")
        print(f"Backup Interval: Every {BACKUP_CYCLE_INTERVAL} cycles")
    print(f"API Posting: {'ENABLED' if ENABLE_API_POSTING else 'DISABLED'}")
    if ENABLE_API_POSTING:
        print(f"API URL: {API_URL}")
    print("=" * 100)
    
    date_folder = create_folder()
    excel_path = get_excel_path(date_folder) if date_folder else None
    
    print("\nInitializing Chrome driver...")
    driver = setup_driver()
    print("✓ Browser opened")
    print("=" * 100)
    
    # Navigate to page
    print(f"\nNavigating to: {URL}")
    driver.get(URL)
    time.sleep(3)
    
    # Validate CAPTCHA
    if not validate_captcha(driver):
        print("\n✗ CAPTCHA validation failed. Exiting...")
        driver.quit()
        return
    
    # Wait for display board to load
    time.sleep(3)
    
    cycle_count = 0
    first_cycle = True
    last_backup_cycle = 0
    
    try:
        while True:
            cycle_count += 1
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            if ENABLE_EXCEL_SAVING and date_folder:
                current_date_folder = get_date_folder()
                if current_date_folder != date_folder:
                    print(f"\n{'='*100}")
                    print(f"📅 DATE CHANGED - NEW DAY STARTED")
                    print(f"{'='*100}\n")
                    
                    date_folder = create_folder()
                    excel_path = get_excel_path(date_folder)
                    first_cycle = True
                    last_backup_cycle = 0
                    cycle_count = 1
            
            print(f"\n{'='*100}")
            print(f"CYCLE {cycle_count} - {current_time}")
            print(f"{'='*100}")
            
            courts_data = scrape_display_board(driver)
            
            if courts_data:
                excel_success = False
                api_result = None
                
                if ENABLE_EXCEL_SAVING and excel_path:
                    excel_success = save_to_excel(courts_data, excel_path, open_file=first_cycle)
                
                if ENABLE_API_POSTING:
                    api_result = post_all_courts_to_api(courts_data)
                
                print(f"\n{'='*100}")
                print(f"✓✓✓ CYCLE {cycle_count} COMPLETED ✓✓✓")
                print(f"   Extracted: {len(courts_data)} records")
                
                if ENABLE_EXCEL_SAVING:
                    status = "SUCCESS" if excel_success else "FAILED"
                    print(f"   Excel Save: {status}")
                
                if ENABLE_API_POSTING and api_result:
                    print(f"   API Posting: {api_result['successful']}/{api_result['total']} successful")
                
                print(f"{'='*100}")
                
                if excel_success:
                    first_cycle = False
                    
                    if ENABLE_EXCEL_SAVING and cycle_count - last_backup_cycle >= BACKUP_CYCLE_INTERVAL:
                        backup_success = create_backup_from_main_excel(excel_path, date_folder)
                        if backup_success:
                            last_backup_cycle = cycle_count
            else:
                print(f"\n   ✗ No data scraped in cycle {cycle_count}")
            
            next_time = datetime.fromtimestamp(time.time() + SCRAPE_INTERVAL).strftime('%Y-%m-%d %H:%M:%S')
            print(f"\n{'─'*100}")
            print(f"⏳ Waiting {SCRAPE_INTERVAL} seconds | Next cycle: {next_time}")
            if ENABLE_EXCEL_SAVING:
                cycles_until_backup = BACKUP_CYCLE_INTERVAL - (cycle_count - last_backup_cycle)
                print(f"   Next backup in: {cycles_until_backup} cycle(s)")
            print(f"{'─'*100}")
            time.sleep(SCRAPE_INTERVAL)
    
    except KeyboardInterrupt:
        print("\n" + "=" * 100)
        print("⚠ Script stopped by user")
        print(f"Total cycles completed: {cycle_count}")
        print("=" * 100)
    
    except Exception as e:
        print(f"\n✗ Unexpected error: {str(e)}")
        import traceback
        traceback.print_exc()
    
    finally:
        print("\nClosing browser...")
        driver.quit()
        print("✓ Script terminated")


if __name__ == "__main__":
    main()