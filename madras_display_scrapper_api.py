"""
Madras High Court Detailed Display Board Scraper
URL: https://hcmadras.tn.gov.in/display_board_mhc.php
Extracts: Court No, Item No, Case Number, Case Number (Full)
WITH TIMESTAMPED BACKUP FILES EVERY 60 CYCLES + API INTEGRATION
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
import pandas as pd
import platform
from bs4 import BeautifulSoup
import requests
import json

# ==================== CONFIGURATION ====================
URL = "https://hcmadras.tn.gov.in/display_board_mhc.php"
SCRAPE_INTERVAL = 30  # seconds
BASE_FOLDER = r"D:\CourtDisplayBoardScraper\displayboardexcel\madras_hc_detailed_excel"
BACKUP_CYCLE_INTERVAL = 60  # Create backup after every 60 cycles
BENCH_NAME = "Chennai"

# API Configuration
API_URL = "https://api.courtlivestream.com/api/display-boards/create"
API_TIMEOUT = 10  # seconds
ENABLE_API_POSTING = True  # Set to False to disable API posting
ENABLE_EXCEL_SAVING = True  # Set to False to disable Excel saving

# ==================== HELPER FUNCTIONS ====================

def extract_case_number_numeric(case_full):
    """
    Extract only the numeric part from full case number
    Examples:
    - "WP.1083/2026" -> "1083"
    - "WP1594/2026" -> "1594"
    - "CRL OP.16466/2007" -> "16466"
    """
    try:
        if not case_full or not case_full.strip():
            return ""
        
        # Pattern: Find number before the slash
        match = re.search(r'\.?(\d+)/', case_full)
        if match:
            return match.group(1).strip()
        
        # Fallback: Find any number
        match = re.search(r'\b(\d+)\b', case_full)
        if match:
            return match.group(1).strip()
        
        return ""
    except:
        return ""


def extract_item_number_numeric(item_str):
    """
    Extract only the numeric part from item number
    Examples:
    - "2" -> "2"
    - "6/L1" -> "6"
    - "21" -> "21"
    """
    try:
        if not item_str or not item_str.strip():
            return ""
        
        # Pattern: Find number before slash or standalone
        match = re.search(r'^(\d+)', item_str)
        if match:
            return match.group(1).strip()
        
        return ""
    except:
        return ""


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
        
        # API MAPPING:
        # Court No -> courtHallNumber
        # Case Number -> caseNumber (numeric only)
        # Item No -> serialNumber (numeric as int)
        
        court_no = court_data.get("Court No", "")
        case_number = court_data.get("Case Number", "")
        item_no = court_data.get("Item No", "")
        
        # Convert item number to integer for serialNumber
        try:
            serial_number = int(item_no) if item_no else 0
        except (ValueError, TypeError):
            serial_number = 0
        
        payload = {
            "benchName": BENCH_NAME,
            "courtHallNumber": court_no,
            "caseNumber": case_number,  # Numeric only
            "serialNumber": serial_number,  # From Item No
            "date": date_str,
            "time": time_str,
            "stage": court_data.get("Case Number (Full)", ""),  # Full case as stage
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
    print(f"MAPPING: Court No->courtHallNumber | Case Number->caseNumber | Item No->serialNumber")
    print(f"{'='*100}\n")
    
    for idx, court_data in enumerate(courts_data_list, 1):
        court_no = court_data.get("Court No", "N/A")
        case_num = court_data.get("Case Number", "N/A")
        item_no = court_data.get("Item No", "N/A")
        
        print(f"   [{idx}/{total_courts}] Court={court_no} | Case={case_num} | Item={item_no}", end=" ")
        
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
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36")
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.implicitly_wait(10)
    return driver


def create_folder():
    """Create date-based folder structure"""
    if not ENABLE_EXCEL_SAVING:
        return None
        
    current_date = datetime.now().strftime("%Y_%m_%d")
    date_folder = os.path.join(BASE_FOLDER, f"madras_hc_detailed_{current_date}")
    
    if not os.path.exists(date_folder):
        os.makedirs(date_folder)
        print(f"✓ Created folder: {date_folder}")
    
    return date_folder


def get_date_folder():
    """Get today's date-based folder path"""
    current_date = datetime.now().strftime("%Y_%m_%d")
    date_folder = os.path.join(BASE_FOLDER, f"madras_hc_detailed_{current_date}")
    return date_folder


def get_excel_path(folder):
    """Get full path for today's main Excel file"""
    if not folder:
        return None
    current_date = datetime.now().strftime("%Y_%m_%d")
    filename = f"madras_hc_detailed_{current_date}.xlsx"
    excel_path = os.path.join(folder, filename)
    return excel_path


def get_timestamped_backup_path(folder):
    """Get full path for timestamped backup Excel file"""
    if not folder:
        return None
    current_timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M")
    filename = f"madras_hc_detailed_bk_{current_timestamp}.xlsx"
    backup_path = os.path.join(folder, filename)
    return backup_path


def create_backup_from_main_excel(main_excel_path, folder):
    """Create a timestamped backup file by copying ALL data from the main Excel file"""
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
        print(f"   Source: {os.path.basename(main_excel_path)}")
        print(f"   Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"{'='*100}\n")
        
        return True
        
    except Exception as e:
        print(f"   ✗ Error creating timestamped backup: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


def open_excel_file(file_path):
    """Open Excel file automatically after first save"""
    try:
        if platform.system() == 'Windows':
            os.file(file_path)
            print(f"   ✓ Excel file opened: {file_path}")
    except Exception as e:
        print(f"   ⚠ Could not auto-open Excel: {str(e)}")


def extract_text_from_element(element):
    """Extract clean text from element"""
    try:
        text = element.text.strip()
        return text
    except:
        return ""


# ==================== SCRAPING FUNCTIONS ====================

def scrape_display_board(driver):
    """
    Scrape courts from Madras HC detailed display board
    Layout: Grid of court boxes arranged in rows
    - 6 rows with 4 courts each = 24 courts
    - 1 last row with 2 courts = 2 courts
    - Total: 26 courts
    Each box contains: Judge photos, Court No, Item No, Case Number
    """
    try:
        print("   → Loading display board page...")
        driver.get(URL)
        
        # Wait for page to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        time.sleep(5)  # Extra wait for dynamic content
        
        scrape_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        print("\n" + "="*100)
        print("ANALYZING PAGE STRUCTURE - EXTRACTING ALL COURTS...")
        print("="*100)
        
        all_courts_data = []
        
        # Get page source and parse with BeautifulSoup
        page_source = driver.page_source
        soup = BeautifulSoup(page_source, 'html.parser')
        
        # Method 1: Find all text containing "Court No :"
        page_text = soup.get_text()
        
        # Find ALL occurrences of "Court No" to count total courts
        court_no_pattern = r'Court No\s*:\s*(\d+)'
        all_court_nos = re.findall(court_no_pattern, page_text)
        total_courts_found = len(all_court_nos)
        
        print(f"   → Found {total_courts_found} 'Court No :' instances on page")
        
        # Split the page text into sections by "Court No :"
        # This ensures we capture each court individually
        court_sections = re.split(r'Court No\s*:', page_text)
        
        print(f"   → Split into {len(court_sections)} sections")
        
        court_count = 0
        
        # Skip first section (before first "Court No :")
        for section_idx, section in enumerate(court_sections[1:], 1):
            try:
                # Extract Court No from start of section
                court_no_match = re.match(r'\s*(\d+)', section)
                if not court_no_match:
                    continue
                
                court_no = court_no_match.group(1).strip()
                
                # Extract Item No
                item_match = re.search(r'Item No\s*:\s*([\d/L]+)', section)
                if not item_match:
                    print(f"      ⚠ Section {section_idx}: Court {court_no} - No Item No found")
                    continue
                
                item_no_full = item_match.group(1).strip()
                
                # Extract Case Number - look for patterns like WP.1083/2026, CRL OP.16466/2007
                case_match = re.search(r'((?:WP|CRL|SA|OP|WRIT|MA|CS|OSA|COC)[.\s]*\d+/\d+)', section, re.IGNORECASE)
                if not case_match:
                    # Try simpler pattern for cases without prefix
                    case_match = re.search(r'(\d+/\d+)', section)
                    if not case_match:
                        print(f"      ⚠ Section {section_idx}: Court {court_no} - No Case Number found")
                        continue
                
                case_number_full = case_match.group(1).strip()
                # Clean up case number (remove extra spaces)
                case_number_full = re.sub(r'\s+', '', case_number_full)
                
                # Extract numeric parts
                item_no_numeric = extract_item_number_numeric(item_no_full)
                case_number_numeric = extract_case_number_numeric(case_number_full)
                
                court_data = {
                    "Bench Name": BENCH_NAME,
                    "Court No": court_no,
                    "Item No": item_no_numeric,
                    "Case Number": case_number_numeric,
                    "Case Number (Full)": case_number_full,
                    "DateTime": scrape_time
                }
                
                all_courts_data.append(court_data)
                court_count += 1
                print(f"      ✓ Court {court_no}: Item {item_no_numeric} | Case {case_number_numeric} ({case_number_full})")
                
            except Exception as e:
                print(f"      ✗ Error processing section {section_idx}: {str(e)}")
                continue
        
        print(f"\n{'='*100}")
        print(f"EXTRACTION SUMMARY:")
        print(f"{'='*100}")
        print(f"   ✓ Total courts extracted: {len(all_courts_data)}")
        print(f"   ✓ Expected: 26 courts (6 rows × 4 + 1 row × 2)")
        print(f"   ✓ Timestamp: {scrape_time}")
        
        if len(all_courts_data) < 26:
            print(f"   ⚠ WARNING: Expected 26 courts but extracted {len(all_courts_data)}")
        
        if all_courts_data:
            print(f"\n   Sample extracted data (first 5 courts):")
            sample_size = min(5, len(all_courts_data))
            for i, court in enumerate(all_courts_data[:sample_size], 1):
                print(f"      {i}. Court {court['Court No']}: Item {court['Item No']} | Case {court['Case Number']} ({court['Case Number (Full)']})")
            
            if len(all_courts_data) > 5:
                print(f"\n   Last court extracted:")
                last = all_courts_data[-1]
                print(f"      Court {last['Court No']}: Item {last['Item No']} | Case {last['Case Number']} ({last['Case Number (Full)']})")
        
        print(f"{'='*100}\n")
        
        return all_courts_data
    
    except Exception as e:
        print(f"\n   ✗ ERROR during scraping: {str(e)}")
        import traceback
        traceback.print_exc()
        return []


# ==================== EXCEL SAVE FUNCTIONS ====================

def save_to_excel(data, file_path, open_file=False):
    """Save scraped data to main Excel file"""
    if not ENABLE_EXCEL_SAVING or not file_path:
        print("\n   ⚠ Excel saving is DISABLED in configuration")
        return False
        
    try:
        if not data:
            print("   ⚠ No data to save")
            return False
        
        df = pd.DataFrame(data)
        df = df[["Bench Name", "Court No", "Item No", "Case Number", "Case Number (Full)", "DateTime"]]
        
        if os.path.exists(file_path):
            existing_df = pd.read_excel(file_path, engine='openpyxl')
            combined_df = pd.concat([existing_df, df], ignore_index=True)
            combined_df.to_excel(file_path, index=False, sheet_name='Sheet1', engine='openpyxl')
            
            print(f"\n   ✓ Data appended to Excel")
            print(f"   ✓ Added {len(df)} courts (Total: {len(combined_df)} rows)")
            print(f"   ✓ File: {os.path.basename(file_path)}")
        else:
            df.to_excel(file_path, index=False, sheet_name='Sheet1', engine='openpyxl')
            print(f"\n   ✓ New Excel file created")
            print(f"   ✓ Initial data: {len(df)} courts")
            print(f"   ✓ File: {os.path.basename(file_path)}")
        
        if open_file:
            open_excel_file(file_path)
        
        return True
        
    except Exception as e:
        print(f"   ✗ Error saving to Excel: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


# ==================== MAIN EXECUTION ====================

def main():
    """Main execution"""
    print("=" * 100)
    print(" " * 20 + "MADRAS HIGH COURT DETAILED DISPLAY BOARD SCRAPER")
    print(" " * 20 + "WITH EXCEL BACKUP + API INTEGRATION")
    print("=" * 100)
    print(f"URL: {URL}")
    print(f"Scrape Interval: {SCRAPE_INTERVAL} seconds")
    print(f"Excel Columns: Bench Name | Court No | Item No | Case Number | Case Number (Full)")
    print(f"Excel Saving: {'ENABLED' if ENABLE_EXCEL_SAVING else 'DISABLED'}")
    if ENABLE_EXCEL_SAVING:
        print(f"Base Location: {BASE_FOLDER}")
        print(f"Backup Interval: Every {BACKUP_CYCLE_INTERVAL} cycles")
    print(f"API Posting: {'ENABLED' if ENABLE_API_POSTING else 'DISABLED'}")
    if ENABLE_API_POSTING:
        print(f"API URL: {API_URL}")
        print(f"   Court No -> courtHallNumber")
        print(f"   Case Number (numeric) -> caseNumber")
        print(f"   Item No (numeric) -> serialNumber")
    print(f"Bench Name: {BENCH_NAME}")
    print("=" * 100)
    
    date_folder = create_folder()
    excel_path = get_excel_path(date_folder) if date_folder else None
    
    if ENABLE_EXCEL_SAVING and excel_path:
        print(f"✓ Today's folder: {os.path.basename(date_folder)}")
        print(f"✓ Main Excel file: {os.path.basename(excel_path)}")
    print("=" * 100)
    
    print("\nInitializing Chrome driver...")
    driver = setup_driver()
    print("✓ Browser opened")
    print("=" * 100)
    
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
                print(f"   Extracted: {len(courts_data)} courts")
                
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