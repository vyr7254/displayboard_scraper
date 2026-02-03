"""
Patna High Court Display Board Scraper
Extracts court data from Patna HC display board and saves to Excel
WITH TIMESTAMPED BACKUP FILES EVERY 60 CYCLES + API INTEGRATION
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
import requests
import json

# ==================== CONFIGURATION ====================
URL = "https://patnahighcourt.gov.in/online_display_board"
SCRAPE_INTERVAL = 30  # seconds
BASE_FOLDER = r"D:\CourtDisplayBoardScraper\displayboardexcel\patna_hc_excel\patna_bench"
BACKUP_CYCLE_INTERVAL = 60  # Create backup after every 60 cycles
SUB_BENCH_NO = "34"  # Sub-bench number for Patna
BENCH_NAME = "patna"

# API Configuration
API_URL = "https://api.courtlivestream.com/api/display-boards/create"
API_TIMEOUT = 10  # seconds
ENABLE_API_POSTING = True  # Set to False to disable API posting
ENABLE_EXCEL_SAVING = True  # Set to False to disable Excel saving

# ==================== API FUNCTIONS ====================

def post_court_data_to_api(court_data):
    """Post a single court record to the API"""
    try:
        # Extract date and time from DateTime field
        datetime_str = court_data.get("DateTime", "")
        
        if datetime_str:
            try:
                dt_obj = datetime.strptime(datetime_str, "%Y-%m-%d %H:%M:%S")
                date_str = dt_obj.strftime("%Y-%m-%d")  # yyyy-mm-dd format
                time_str = dt_obj.strftime("%I:%M %p")
            except ValueError:
                date_str = datetime.now().strftime("%Y-%m-%d")
                time_str = datetime.now().strftime("%I:%M %p")
        else:
            date_str = datetime.now().strftime("%Y-%m-%d")
            time_str = datetime.now().strftime("%I:%M %p")
        
        # Convert Item No to integer (this serves as serialNumber)
        serial_number = court_data.get("Item No", "")
        try:
            serial_number = int(serial_number) if serial_number else 0
        except (ValueError, TypeError):
            serial_number = 0
        
        # Prepare API payload with mapped fields
        # Mapping: benchName:Bench Name, courtHallNumber:Court Number, caseNumber:Case Number, serialNumber:Item No
        payload = {
            "benchName": court_data.get("Bench Name", ""),
            "courtHallNumber": court_data.get("Court Number", ""),
            "caseNumber": court_data.get("Case Number", ""),
            "serialNumber": serial_number,
            "date": date_str,
            "time": time_str
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
        court_no = court_data.get("Court Number", "N/A")
        case_no = court_data.get("Case Number", "N/A")
        
        print(f"   [{idx}/{total_courts}] Court {court_no} (Item: {court_data.get('Item No', 'N/A')})...", end=" ")
        
        success, response = post_court_data_to_api(court_data)
        
        if success:
            successful_posts += 1
            print("✓")
        else:
            failed_posts += 1
            print(f"✗ ({response})")
            errors.append({"court": court_no, "case": case_no, "error": response})
    
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


# ==================== SETUP FUNCTIONS ===================
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
    """
    Create date-based folder structure with bench folder
    Format: D:\CourtDisplayBoardScraper\displayboardexcel\patna_hc_excel\patna_bench\patna_YYYY_MM_DD\
    """
    if not ENABLE_EXCEL_SAVING:
        return None
        
    if not ENABLE_EXCEL_SAVING:
        return None
        
    current_date = datetime.now().strftime("%Y_%m_%d")
    date_folder = os.path.join(BASE_FOLDER, f"patna_{current_date}")
    
    if not os.path.exists(date_folder):
        os.makedirs(date_folder)
        print(f"✓ Created folder: {date_folder}")
    
    return date_folder


def get_date_folder():
    """Get today's date-based folder path"""
    current_date = datetime.now().strftime("%Y_%m_%d")
    date_folder = os.path.join(BASE_FOLDER, f"patna_{current_date}")
    return date_folder


def get_excel_path(folder):
    """
    Get full path for today's main Excel file
    Format: patna_YYYY_MM_DD.xlsx
    """
    if not folder:
        return None
    if not folder:
        return None
    current_date = datetime.now().strftime("%Y_%m_%d")
    filename = f"patna_{current_date}.xlsx"
    excel_path = os.path.join(folder, filename)
    return excel_path


def get_timestamped_backup_path(folder):
    """
    Get full path for timestamped backup Excel file
    Format: patna_bk_YYYY_MM_DD_HH_MM.xlsx
    """
    if not folder:
        return None
    if not folder:
        return None
    current_timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M")
    filename = f"patna_bk_{current_timestamp}.xlsx"
    backup_path = os.path.join(folder, filename)
    return backup_path


def create_backup_from_main_excel(main_excel_path, folder):
    """
    Create a timestamped backup file by copying ALL data from the main Excel file
    This ensures we have a complete snapshot at the time of backup
    """
    if not ENABLE_EXCEL_SAVING or not main_excel_path or not folder:
        return False
        
    if not ENABLE_EXCEL_SAVING or not main_excel_path or not folder:
        return False
        
    try:
        # Check if main Excel file exists
        if not os.path.exists(main_excel_path):
            print("   ⚠ Main Excel file does not exist yet. Cannot create backup.")
            return False
        
        # Read all data from main Excel file
        main_df = pd.read_excel(main_excel_path, engine='openpyxl')
        
        if main_df.empty:
            print("   ⚠ Main Excel file is empty. No backup created.")
            return False
        
        # Generate timestamped backup path
        backup_path = get_timestamped_backup_path(folder)
        
        # Save complete data to new backup file
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
            os.startfile(file_path)
            print(f"   ✓ Excel file opened: {file_path}")
    except Exception as e:
        print(f"   ⚠ Could not auto-open Excel: {str(e)}")


def extract_cell_content(cell):
    """
    Extract content from cell - either input button value or text content
    Uses innerHTML parsing for speed
    """
    try:
        # Get the inner HTML once
        html_content = cell.get_attribute('innerHTML')
        
        # Parse with BeautifulSoup
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Check for input button first
        input_tag = soup.find('input')
        if input_tag and input_tag.get('value'):
            return input_tag.get('value').strip()
        
        # Otherwise get text content
        text = soup.get_text(separator=' ', strip=True)
        
        # Clean up extra whitespace
        text = re.sub(r'\s+', ' ', text).strip()
        
        return text
    except:
        return ""


def extract_case_info(case_text):
    """
    Extract Item No and Case No from the case text
    Format: "13 - C.MISC./459/2016 (FOR ADMISSION)"
    Returns: (item_no, case_no)
    """
    if not case_text or case_text == "NOT IN SESSION":
        return "", ""
    
    try:
        # Split by first dash to separate item number and case details
        if ' - ' in case_text:
            parts = case_text.split(' - ', 1)
            item_no = parts[0].strip()
            case_no = parts[1].strip() if len(parts) > 1 else ""
            return item_no, case_no
        else:
            # If no dash, entire text is case number
            return "", case_text.strip()
    except:
        return "", case_text


# ==================== SCRAPING FUNCTIONS ====================

def scrape_display_board(driver):
    """
    Scrape courts from Patna High Court display board
    Table structure: 4 columns arranged as:
    [Court Number 1] [Case Number 1] [Court Number 2] [Case Number 2]
    Each row contains 2 courts
    """
    try:
        print("   → Loading display board page...")
        driver.get(URL)
        
        # Wait for table to load
        print("   → Waiting for table to appear...")
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "CSSTableDisplayBoard"))
        )
        print("   → Table found, waiting for content to load...")
        time.sleep(10)  # Longer wait for dynamic content and flip-card animations
        
        # Get current timestamp
        scrape_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        print("\n" + "="*100)
        print("ANALYZING PAGE STRUCTURE - EXTRACTING ALL COURTS...")
        print("="*100)
        
        # Find the display board table
        table = driver.find_element(By.CLASS_NAME, "CSSTableDisplayBoard")
        rows = table.find_elements(By.TAG_NAME, "tr")
        print(f"   → Found {len(rows)} total rows in table")
        
        all_courts_data = []
        
        # Check headers (should be in first data row after initial header)
        print(f"\n{'─'*100}")
        print("ANALYZING TABLE STRUCTURE...")
        print(f"{'─'*100}")
        
        # Find the header row (contains "COURT NUMBER" and "CASE NUMBER")
        header_row_idx = 0
        for idx, row in enumerate(rows):
            cells = row.find_elements(By.TAG_NAME, "th")
            if not cells:
                cells = row.find_elements(By.TAG_NAME, "td")
            
            cell_texts = [extract_cell_content(cell) for cell in cells]
            if any("COURT NUMBER" in text.upper() for text in cell_texts):
                header_row_idx = idx
                print(f"   → Found header at row {idx}")
                break
        
        print(f"\n{'─'*100}")
        print("EXTRACTING DATA...")
        print(f"{'─'*100}")
        
        # Process data rows (after header)
        print(f"   → Processing data rows...")
        
        for row_idx in range(header_row_idx + 1, len(rows)):
            row = rows[row_idx]
            
            try:
                # Get all cells (td elements) at once
                cells = row.find_elements(By.TAG_NAME, "td")
                
                # Skip rows with insufficient cells
                if len(cells) < 2:
                    continue
                
                # Process first court (columns 0 and 1)
                if len(cells) >= 2:
                    court_num_1 = extract_cell_content(cells[0])
                    case_details_1 = extract_cell_content(cells[1])
                    
                    if court_num_1:
                        item_no_1, case_no_1 = extract_case_info(case_details_1)
                        
                        all_courts_data.append({
                            "Bench Name": BENCH_NAME,
                            "SubBenchNo": SUB_BENCH_NO,
                            "Court Number": court_num_1,
                            "Item No": item_no_1,
                            "Case Number": case_no_1,
                            "Full Case Details": case_details_1,
                            "DateTime": scrape_time
                        })
                
                # Process second court (columns 2 and 3)
                if len(cells) >= 4:
                    court_num_2 = extract_cell_content(cells[2])
                    case_details_2 = extract_cell_content(cells[3])
                    
                    if court_num_2:
                        item_no_2, case_no_2 = extract_case_info(case_details_2)
                        
                        all_courts_data.append({
                            "Bench Name": BENCH_NAME,
                            "SubBenchNo": SUB_BENCH_NO,
                            "Court Number": court_num_2,
                            "Item No": item_no_2,
                            "Case Number": case_no_2,
                            "Full Case Details": case_details_2,
                            "DateTime": scrape_time
                        })
                
            except Exception as e:
                continue
        
        print(f"\n{'='*100}")
        print(f"EXTRACTION SUMMARY:")
        print(f"{'='*100}")
        print(f"   ✓ Total courts extracted: {len(all_courts_data)}")
        print(f"   ✓ Timestamp: {scrape_time}")
        
        if all_courts_data:
            print(f"\n   Sample extracted data:")
            sample_size = min(5, len(all_courts_data))
            for i, court in enumerate(all_courts_data[:sample_size], 1):
                if court['Case Number']:
                    print(f"      {i}. Court {court['Court Number']} | Item {court['Item No']} | {court['Case Number']}")
                else:
                    print(f"      {i}. Court {court['Court Number']} | {court['Full Case Details']}")
        
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
        
    if not ENABLE_EXCEL_SAVING or not file_path:
        print("\n   ⚠ Excel saving is DISABLED in configuration")
        return False
        
    try:
        if not data:
            print("   ⚠ No data to save")
            return False
        
        df = pd.DataFrame(data)
        df = df[["Bench Name", "SubBenchNo", "Court Number", "Item No", "Case Number", "Full Case Details", "DateTime"]]
        
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
    print(" " * 25 + "PATNA HIGH COURT - PATNA BENCH DISPLAY BOARD SCRAPER")
    print(" " * 25 + "WITH EXCEL BACKUP + API INTEGRATION")
    print(" " * 25 + "WITH EXCEL BACKUP + API INTEGRATION")
    print("=" * 100)
    print(f"URL: {URL}")
    print(f"Scrape Interval: {SCRAPE_INTERVAL} seconds")
    print(f"Excel Saving: {'ENABLED' if ENABLE_EXCEL_SAVING else 'DISABLED'}")
    if ENABLE_EXCEL_SAVING:
        print(f"Base Location: {BASE_FOLDER}")
        print(f"Backup Interval: Every {BACKUP_CYCLE_INTERVAL} cycles")
    print(f"API Posting: {'ENABLED' if ENABLE_API_POSTING else 'DISABLED'}")
    if ENABLE_API_POSTING:
        print(f"API URL: {API_URL}")
    print(f"Excel Saving: {'ENABLED' if ENABLE_EXCEL_SAVING else 'DISABLED'}")
    if ENABLE_EXCEL_SAVING:
        print(f"Base Location: {BASE_FOLDER}")
        print(f"Backup Interval: Every {BACKUP_CYCLE_INTERVAL} cycles")
    print(f"API Posting: {'ENABLED' if ENABLE_API_POSTING else 'DISABLED'}")
    if ENABLE_API_POSTING:
        print(f"API URL: {API_URL}")
    print(f"SubBench Number: {SUB_BENCH_NO} (applied to all records)")
    print(f"Bench Name: {BENCH_NAME} (applied to all records)")
    print("=" * 100)
    
    # Get today's folder and file paths
    date_folder = create_folder()
    excel_path = get_excel_path(date_folder) if date_folder else None
    excel_path = get_excel_path(date_folder) if date_folder else None
    
    if ENABLE_EXCEL_SAVING and excel_path:
        print(f"✓ Today's folder: {os.path.basename(date_folder)}")
        print(f"✓ Main Excel file: {os.path.basename(excel_path)}")
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
            
            # Check if date has changed (new day started) - only if Excel is enabled
            if ENABLE_EXCEL_SAVING and date_folder:
                current_date_folder = get_date_folder()
                if current_date_folder != date_folder:
                    print(f"\n{'='*100}")
                    print(f"📅 DATE CHANGED - NEW DAY STARTED")
                    print(f"   Old folder: {os.path.basename(date_folder)}")
                    print(f"   New folder: {os.path.basename(current_date_folder)}")
                    print(f"{'='*100}\n")
                    
                    # Create new folder and update paths
                    date_folder = create_folder()
                    excel_path = get_excel_path(date_folder)
                    first_cycle = True
                    last_backup_cycle = 0
                    cycle_count = 1  # Reset cycle count for new day
                    
                    print(f"✓ New main file: {os.path.basename(excel_path)}")
            # Check if date has changed (new day started) - only if Excel is enabled
            if ENABLE_EXCEL_SAVING and date_folder:
                current_date_folder = get_date_folder()
                if current_date_folder != date_folder:
                    print(f"\n{'='*100}")
                    print(f"📅 DATE CHANGED - NEW DAY STARTED")
                    print(f"   Old folder: {os.path.basename(date_folder)}")
                    print(f"   New folder: {os.path.basename(current_date_folder)}")
                    print(f"{'='*100}\n")
                    
                    # Create new folder and update paths
                    date_folder = create_folder()
                    excel_path = get_excel_path(date_folder)
                    first_cycle = True
                    last_backup_cycle = 0
                    cycle_count = 1  # Reset cycle count for new day
                    
                    print(f"✓ New main file: {os.path.basename(excel_path)}")
            
            print(f"\n{'='*100}")
            print(f"CYCLE {cycle_count} - {current_time}")
            if ENABLE_EXCEL_SAVING and excel_path:
                print(f"Folder: {os.path.basename(date_folder)}")
                print(f"Main Excel: {os.path.basename(excel_path)}")
            if ENABLE_EXCEL_SAVING and excel_path:
                print(f"Folder: {os.path.basename(date_folder)}")
                print(f"Main Excel: {os.path.basename(excel_path)}")
            print(f"{'='*100}")
            
            courts_data = scrape_display_board(driver)
            
            if courts_data:
                excel_success = False
                api_result = None
                
                # Save to Excel if enabled
                if ENABLE_EXCEL_SAVING and excel_path:
                    excel_success = save_to_excel(courts_data, excel_path, open_file=first_cycle)
                excel_success = False
                api_result = None
                
                # Save to Excel if enabled
                if ENABLE_EXCEL_SAVING and excel_path:
                    excel_success = save_to_excel(courts_data, excel_path, open_file=first_cycle)
                
                # Post to API if enabled
                if ENABLE_API_POSTING:
                    api_result = post_all_courts_to_api(courts_data)
                
                print(f"\n{'='*100}")
                print(f"✓✓✓ CYCLE {cycle_count} COMPLETED ✓✓✓")
                print(f"   Extracted: {len(courts_data)} courts from Patna Bench")
                
                if ENABLE_EXCEL_SAVING:
                    status = "SUCCESS" if excel_success else "FAILED"
                    print(f"   Excel Save: {status}")
                
                if ENABLE_API_POSTING and api_result:
                    print(f"   API Posting: {api_result['successful']}/{api_result['total']} successful")
                
                print(f"{'='*100}")
                
                if excel_success:
                # Post to API if enabled
                if ENABLE_API_POSTING:
                    api_result = post_all_courts_to_api(courts_data)
                
                print(f"\n{'='*100}")
                print(f"✓✓✓ CYCLE {cycle_count} COMPLETED ✓✓✓")
                print(f"   Extracted: {len(courts_data)} courts from Patna Bench")
                
                if ENABLE_EXCEL_SAVING:
                    status = "SUCCESS" if excel_success else "FAILED"
                    print(f"   Excel Save: {status}")
                
                if ENABLE_API_POSTING and api_result:
                    print(f"   API Posting: {api_result['successful']}/{api_result['total']} successful")
                
                print(f"{'='*100}")
                
                if excel_success:
                    first_cycle = False
                    
                    # Check if backup is needed (every 60 cycles)
                    if ENABLE_EXCEL_SAVING and cycle_count - last_backup_cycle >= BACKUP_CYCLE_INTERVAL:
                    if ENABLE_EXCEL_SAVING and cycle_count - last_backup_cycle >= BACKUP_CYCLE_INTERVAL:
                        print(f"\n{'─'*100}")
                        print(f"⏰ BACKUP TIME - {BACKUP_CYCLE_INTERVAL} cycles completed")
                        print(f"   Creating timestamped backup from main Excel file")
                        print(f"{'─'*100}")
                        
                        backup_success = create_backup_from_main_excel(excel_path, date_folder)
                        
                        if backup_success:
                            last_backup_cycle = cycle_count
                            print(f"   ✓ Backup created successfully")
                            print(f"   ✓ This backup contains all data up to cycle {cycle_count}")
            else:
                print(f"\n   ✗ No data scraped in cycle {cycle_count}")
            
            next_time = datetime.fromtimestamp(time.time() + SCRAPE_INTERVAL).strftime('%Y-%m-%d %H:%M:%S')
            
            print(f"\n{'─'*100}")
            print(f"⏳ Waiting {SCRAPE_INTERVAL} seconds")
            print(f"   Next cycle: {next_time}")
            if ENABLE_EXCEL_SAVING:
                cycles_until_backup = BACKUP_CYCLE_INTERVAL - (cycle_count - last_backup_cycle)
                print(f"   Next backup in: {cycles_until_backup} cycle(s)")
            if ENABLE_EXCEL_SAVING:
                cycles_until_backup = BACKUP_CYCLE_INTERVAL - (cycle_count - last_backup_cycle)
                print(f"   Next backup in: {cycles_until_backup} cycle(s)")
            print(f"{'─'*100}")
            time.sleep(SCRAPE_INTERVAL)
    
    except KeyboardInterrupt:
        print("\n" + "=" * 100)
        print("⚠ Script stopped by user")
        print(f"Total cycles completed: {cycle_count}")
        if ENABLE_EXCEL_SAVING and date_folder and excel_path:
            print(f"Final folder: {os.path.basename(date_folder)}")
            print(f"Final main file: {os.path.basename(excel_path)}")
        if ENABLE_EXCEL_SAVING and date_folder and excel_path:
            print(f"Final folder: {os.path.basename(date_folder)}")
            print(f"Final main file: {os.path.basename(excel_path)}")
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