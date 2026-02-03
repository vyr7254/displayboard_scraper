"""
Delhi High Court Display Board Scraper
URL: https://delhihighcourt.nic.in/app/physical-display-board
Extracts: Court, Item No, Hon'ble Judges, Case No, Case No (Full), Title
WITH TIMESTAMPED BACKUP FILES EVERY 60 CYCLES + API INTEGRATION
Scrapes every 30 seconds with automatic page refresh
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
from selenium.webdriver.support.ui import Select
import pandas as pd
import platform
from bs4 import BeautifulSoup
import requests
import json

# ==================== CONFIGURATION ====================
URL = "https://delhihighcourt.nic.in/app/physical-display-board"
SCRAPE_INTERVAL = 30  # seconds
BASE_FOLDER = r"D:\CourtDisplayBoardScraper\displayboardexcel\delhi_hc_excel"
BACKUP_CYCLE_INTERVAL = 60  # Create backup after every 60 cycles
BENCH_NAME = "New Delhi"

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
    - "LPA - 500 / 2025" -> "500"
    - "CONT.CAS(C) - 155 / 2026" -> "155"
    - "W.P.(C) - 696 / 2025" -> "696"
    """
    try:
        if not case_full or not case_full.strip():
            return ""
        
        # Pattern: Find number before the slash
        match = re.search(r'-\s*(\d+)\s*/', case_full)
        if match:
            return match.group(1).strip()
        
        # Fallback: Find any number
        match = re.search(r'\b(\d+)\b', case_full)
        if match:
            return match.group(1).strip()
        
        return ""
    except:
        return ""


def split_title_petitioner_respondent(title_str):
    """
    Split title into petitioner and respondent based on 'Vs' or 'vs'
    Example:
    "NETAJI SUBHAS UNIVERSITY Vs SH DHRUW KANT JHA & ORS."
    -> petitioner: "NETAJI SUBHAS UNIVERSITY"
    -> respondent: "SH DHRUW KANT JHA & ORS."
    """
    try:
        if not title_str or not title_str.strip():
            return "", ""
        
        # Split by Vs or vs (case insensitive)
        parts = re.split(r'\s+[Vv][Ss]\s+', title_str, maxsplit=1)
        
        if len(parts) == 2:
            petitioner = parts[0].strip()
            respondent = parts[1].strip()
            return petitioner, respondent
        else:
            # No Vs found, treat entire title as petitioner
            return title_str.strip(), ""
    except:
        return "", ""


def extract_item_number_numeric(item_str):
    """
    Extract numeric part from item number
    Examples:
    - "A23" -> "23"
    - "O50" -> "50"
    - "*" -> ""
    """
    try:
        if not item_str or not item_str.strip() or item_str == "*":
            return ""
        
        # Extract digits from item
        match = re.search(r'(\d+)', item_str)
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
        # Court -> courtHallNumber
        # Case No. -> caseNumber (numeric only)
        # Item No. -> serialNumber (numeric as int)
        # Hon'ble Judges -> judgeName
        # Petitioner -> petitioner
        # Respondent -> respondent
        
        court_no = court_data.get("Court", "")
        case_number = court_data.get("Case No.", "")
        item_no = court_data.get("Item No.", "")
        judges = court_data.get("Hon'ble Judges", "")
        petitioner = court_data.get("Petitioner", "")
        respondent = court_data.get("Respondent", "")
        
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
            "judgeName": judges,  # Hon'ble Judges
            "petitioner": petitioner,
            "respondent": respondent,
            "stage": court_data.get("Title", ""),  # Full title as stage
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
    print(f"{'='*100}\n")
    
    for idx, court_data in enumerate(courts_data_list, 1):
        court_no = court_data.get("Court", "N/A")
        case_num = court_data.get("Case No.", "N/A")
        
        print(f"   [{idx}/{total_courts}] Court={court_no} | Case={case_num}", end=" ")
        
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
    date_folder = os.path.join(BASE_FOLDER, f"delhi_hc_{current_date}")
    
    if not os.path.exists(date_folder):
        os.makedirs(date_folder)
        print(f"✓ Created folder: {date_folder}")
    
    return date_folder


def get_date_folder():
    """Get today's date-based folder path"""
    current_date = datetime.now().strftime("%Y_%m_%d")
    date_folder = os.path.join(BASE_FOLDER, f"delhi_hc_{current_date}")
    return date_folder


def get_excel_path(folder):
    """Get full path for today's main Excel file"""
    if not folder:
        return None
    current_date = datetime.now().strftime("%Y_%m_%d")
    filename = f"delhi_hc_{current_date}.xlsx"
    excel_path = os.path.join(folder, filename)
    return excel_path


def get_timestamped_backup_path(folder):
    """Get full path for timestamped backup Excel file"""
    if not folder:
        return None
    current_timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M")
    filename = f"delhi_hc_bk_{current_timestamp}.xlsx"
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
        print(f"   Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"{'='*100}\n")
        
        return True
        
    except Exception as e:
        print(f"   ✗ Error creating timestamped backup: {str(e)}")
        return False


def open_excel_file(file_path):
    """Open Excel file automatically after first save"""
    try:
        if platform.system() == 'Windows':
            os.startfile(file_path)
            print(f"   ✓ Excel file opened: {file_path}")
    except Exception as e:
        print(f"   ⚠ Could not auto-open Excel: {str(e)}")


# ==================== SCRAPING FUNCTIONS ====================

def scrape_display_board(driver):
    """
    Scrape courts from Delhi High Court display board
    Table with columns: Court | Item No. | Hon'ble Judges | Case No. | Title | VC Link
    """
    try:
        print("   → Loading display board page...")
        driver.get(URL)
        
        # Wait for table to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "physical_display_board"))
        )
        time.sleep(3)
        
        # Select 100 entries per page from dropdown
        try:
            print("   → Selecting 100 entries per page...")
            length_select = Select(driver.find_element(By.NAME, "physical_display_board_length"))
            length_select.select_by_value("100")
            time.sleep(3)  # Wait for table to reload
            print("   ✓ Selected 100 entries per page")
        except Exception as e:
            print(f"   ⚠ Could not change page length: {str(e)}")
        
        scrape_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        print("\n" + "="*100)
        print("EXTRACTING ALL COURTS FROM TABLE...")
        print("="*100)
        
        all_courts_data = []
        
        # Find the table
        table = driver.find_element(By.ID, "physical_display_board")
        tbody = table.find_element(By.TAG_NAME, "tbody")
        rows = tbody.find_elements(By.TAG_NAME, "tr")
        
        print(f"   → Found {len(rows)} rows in table")
        
        for row_idx, row in enumerate(rows, 1):
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                
                if len(cells) < 5:  # Need at least 5 columns
                    continue
                
                # Extract data from cells
                # Column order: Court | Item No. | Hon'ble Judges | Case No. | Title | VC Link
                court_cell = cells[0]
                item_cell = cells[1]
                judges_cell = cells[2]
                case_cell = cells[3]
                title_cell = cells[4]
                
                # Extract text
                court_no = court_cell.text.strip()
                item_no_full = item_cell.text.strip()
                judges = judges_cell.text.strip()
                case_no_full = case_cell.text.strip()
                title = title_cell.text.strip()
                
                # Skip empty rows (marked with *)
                if item_no_full == "*" or not case_no_full:
                    continue
                
                # Extract numeric parts
                item_no_numeric = extract_item_number_numeric(item_no_full)
                case_no_numeric = extract_case_number_numeric(case_no_full)
                
                # Split title into petitioner and respondent
                petitioner, respondent = split_title_petitioner_respondent(title)
                
                court_data = {
                    "Court": court_no,
                    "Item No.": item_no_numeric,
                    "Hon'ble Judges": judges,
                    "Case No.": case_no_numeric,
                    "Case No. (Full)": case_no_full,
                    "Title": title,
                    "Petitioner": petitioner,
                    "Respondent": respondent,
                    "DateTime": scrape_time
                }
                
                all_courts_data.append(court_data)
                print(f"      ✓ Court {court_no}: Item {item_no_numeric} | Case {case_no_numeric}")
                
            except Exception as e:
                print(f"      ✗ Error at row {row_idx}: {str(e)}")
                continue
        
        print(f"\n{'='*100}")
        print(f"EXTRACTION SUMMARY:")
        print(f"{'='*100}")
        print(f"   ✓ Total courts extracted: {len(all_courts_data)}")
        print(f"   ✓ Timestamp: {scrape_time}")
        
        if all_courts_data:
            print(f"\n   Sample extracted data (first 3 courts):")
            sample_size = min(3, len(all_courts_data))
            for i, court in enumerate(all_courts_data[:sample_size], 1):
                print(f"      {i}. Court {court['Court']}: {court['Case No. (Full)']} -> {court['Case No.']}")
        
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
        df = df[["Court", "Item No.", "Hon'ble Judges", "Case No.", "Case No. (Full)", "Title", "Petitioner", "Respondent", "DateTime"]]
        
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
        return False


# ==================== MAIN EXECUTION ====================

def main():
    """Main execution"""
    print("=" * 100)
    print(" " * 20 + "DELHI HIGH COURT DISPLAY BOARD SCRAPER")
    print(" " * 20 + "WITH EXCEL BACKUP + API INTEGRATION")
    print(" " * 20 + "Auto-refresh every 30 seconds")
    print("=" * 100)
    print(f"URL: {URL}")
    print(f"Scrape Interval: {SCRAPE_INTERVAL} seconds")
    print(f"Excel Columns: Court | Item No. | Hon'ble Judges | Case No. | Case No. (Full) | Title | Petitioner | Respondent")
    print(f"Excel Saving: {'ENABLED' if ENABLE_EXCEL_SAVING else 'DISABLED'}")
    if ENABLE_EXCEL_SAVING:
        print(f"Base Location: {BASE_FOLDER}")
        print(f"Backup Interval: Every {BACKUP_CYCLE_INTERVAL} cycles")
    print(f"API Posting: {'ENABLED' if ENABLE_API_POSTING else 'DISABLED'}")
    if ENABLE_API_POSTING:
        print(f"API URL: {API_URL}")
        print(f"   Court -> courtHallNumber")
        print(f"   Case No. (numeric) -> caseNumber")
        print(f"   Item No. (numeric) -> serialNumber")
        print(f"   Hon'ble Judges -> judgeName")
        print(f"   Petitioner -> petitioner")
        print(f"   Respondent -> respondent")
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