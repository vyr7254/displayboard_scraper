"""
Chhattisgarh High Court Display Board Scraper
Extracts court data from Chhattisgarh HC display board and saves to Excel
WITH TIMESTAMPED BACKUP FILES EVERY 60 CYCLES + API INTEGRATION
EXTRACTS: Case Number, Case Type, Case Year from format like "WA / 954 / 2025"
FIXED: Properly extracts data aligned with table headers
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
URL = "https://highcourt.cg.gov.in/hcbspcourtview/court1.php"
SCRAPE_INTERVAL = 30  # seconds
BASE_FOLDER = r"D:\CourtDisplayBoardScraper\displayboardexcel\chhattisgarh_hc_excel"
BACKUP_CYCLE_INTERVAL = 60  # Create backup after every 60 cycles
BENCH_NAME = "Bilaspur"

# API Configuration
API_URL = "https://api.courtlivestream.com/api/display-boards/create"
API_TIMEOUT = 10  # seconds
ENABLE_API_POSTING = True  # Set to False to disable API posting
ENABLE_EXCEL_SAVING = True  # Set to False to disable Excel saving

# ==================== HELPER FUNCTIONS ====================

def parse_case_details(case_str):
    """
    Parse case string to extract case_number, case_type, and case_year
    Examples:
    - "WA / 954 / 2025" -> case_number: "954", case_type: "WA", case_year: "2025"
    - "CRMP / 3873 / 2025" -> case_number: "3873", case_type: "CRMP", case_year: "2025"
    - "FA(MAT) / 448 / 2025" -> case_number: "448", case_type: "FA(MAT)", case_year: "2025"
    - "" or None -> case_number: "", case_type: "", case_year: ""
    """
    try:
        if not case_str or not case_str.strip():
            return "", "", ""
        
        case_str = case_str.strip()
        
        # Pattern: CASE_TYPE / NUMBER / YEAR
        match = re.match(r'([A-Z]+(?:\([A-Z]+\))?)\s*/\s*(\d+)\s*/\s*(\d{4})', case_str)
        
        if match:
            case_type = match.group(1).strip()
            case_number = match.group(2).strip()
            case_year = match.group(3).strip()
            return case_number, case_type, case_year
        
        number_match = re.search(r'/\s*(\d+)\s*/', case_str)
        if number_match:
            return number_match.group(1).strip(), "", ""
        
        return "", "", ""
    
    except Exception as e:
        print(f"   ⚠ Error parsing case details '{case_str}': {str(e)}")
        return "", "", ""


# ==================== API FUNCTIONS ====================

def extract_case_number_from_purpose(purpose_str):
    """
    Extract only the numeric case number from Purpose field
    Examples:
    - "CRMP / 3087 / 2025" -> "3087"
    - "WPS / 8772 / 2025" -> "8772"
    - "CRA / 1349 / 2015" -> "1349"
    """
    try:
        if not purpose_str or not purpose_str.strip():
            return ""
        
        # Pattern: CASE_TYPE / NUMBER / YEAR
        match = re.search(r'/\s*(\d+)\s*/', purpose_str)
        if match:
            return match.group(1).strip()
        
        # Fallback: try to find any number
        match = re.search(r'\b(\d+)\b', purpose_str)
        if match:
            return match.group(1).strip()
        
        return ""
    except:
        return ""


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
        
        # NEW MAPPING:
        # Excel "Purpose" column -> API "caseNumber" (extract numeric part only)
        # Excel "Full Case" column -> API "serialNumber"
        
        # Extract case number from Purpose column (e.g., "CRMP / 3087 / 2025" -> "3087")
        purpose_value = court_data.get("Purpose", "")
        case_number = extract_case_number_from_purpose(purpose_value)
        
        # Get serial number from Full Case column
        full_case_value = court_data.get("Full Case", "")
        try:
            # Handle ranges like "8 - 9" -> take first number
            if isinstance(full_case_value, str) and '-' in full_case_value:
                serial_number = int(full_case_value.split('-')[0].strip())
            else:
                # Try to extract any number from Full Case
                match = re.search(r'\b(\d+)\b', str(full_case_value))
                serial_number = int(match.group(1)) if match else 0
        except (ValueError, TypeError):
            serial_number = 0
        
        list_number = court_data.get("List Type", "")
        try:
            list_number = int(list_number) if list_number else 0
        except (ValueError, TypeError):
            list_number = 0
        
        payload = {
            "benchName": BENCH_NAME,
            "courtHallNumber": court_data.get("Court", ""),
            "caseNumber": case_number,  # From Purpose column (numeric only)
            "serialNumber": serial_number,  # From Full Case column
            "date": date_str,
            "time": time_str,
            "stage": court_data.get("Purpose", ""),
            "listNumber": list_number
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
    print(f"MAPPING: Purpose -> caseNumber (numeric only) | Full Case -> serialNumber")
    print(f"{'='*100}\n")
    
    for idx, court_data in enumerate(courts_data_list, 1):
        court_no = court_data.get("Court", "N/A")
        purpose_val = court_data.get("Purpose", "N/A")
        full_case_val = court_data.get("Full Case", "N/A")
        
        # Extract what will be sent to API
        case_num = extract_case_number_from_purpose(purpose_val)
        
        print(f"   [{idx}/{total_courts}] Court {court_no} | Purpose='{purpose_val}' -> CaseNum={case_num} | FullCase='{full_case_val}' -> SerialNum", end=" ")
        
        success, response = post_court_data_to_api(court_data)
        
        if success:
            successful_posts += 1
            print("✓")
        else:
            failed_posts += 1
            print(f"✗ ({response})")
            errors.append({"court": court_no, "purpose": purpose_val, "error": response})
    
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
            print(f"      - Court {err['court']} (Purpose: {err['purpose']}): {err['error']}")
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
    date_folder = os.path.join(BASE_FOLDER, f"chhattisgarh_hc_{current_date}")
    
    if not os.path.exists(date_folder):
        os.makedirs(date_folder)
        print(f"✓ Created folder: {date_folder}")
    
    return date_folder


def get_date_folder():
    """Get today's date-based folder path"""
    current_date = datetime.now().strftime("%Y_%m_%d")
    date_folder = os.path.join(BASE_FOLDER, f"chhattisgarh_hc_{current_date}")
    return date_folder


def get_excel_path(folder):
    """Get full path for today's main Excel file"""
    if not folder:
        return None
    current_date = datetime.now().strftime("%Y_%m_%d")
    filename = f"chhattisgarh_hc_{current_date}.xlsx"
    excel_path = os.path.join(folder, filename)
    return excel_path


def get_timestamped_backup_path(folder):
    """Get full path for timestamped backup Excel file"""
    if not folder:
        return None
    current_timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M")
    filename = f"chhattisgarh_hc_bk_{current_timestamp}.xlsx"
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
            os.startfile(file_path)
            print(f"   ✓ Excel file opened: {file_path}")
    except Exception as e:
        print(f"   ⚠ Could not auto-open Excel: {str(e)}")


def extract_cell_text(cell):
    """Extract visible text from cell, handling nested HTML elements"""
    try:
        html_content = cell.get_attribute('innerHTML')
        soup = BeautifulSoup(html_content, 'html.parser')
        text = soup.get_text(separator=' ', strip=True)
        text = re.sub(r'\s+', ' ', text).strip()
        return text
    except:
        return ""


# ==================== SCRAPING FUNCTIONS ====================

def scrape_display_board(driver):
    """
    Scrape courts from Chhattisgarh High Court display board
    Uses header mapping to extract data correctly under each column heading
    """
    try:
        print("   → Loading display board page...")
        driver.get(URL)
        
        # Wait for table to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "tb1"))
        )
        time.sleep(5)  # Extra wait for dynamic content
        
        scrape_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        print("\n" + "="*100)
        print("ANALYZING PAGE STRUCTURE - EXTRACTING ALL COURTS...")
        print("="*100)
        
        # Find the main table
        table = driver.find_element(By.ID, "tb1")
        
        # Get headers from thead
        thead = table.find_element(By.TAG_NAME, "thead")
        header_row = thead.find_element(By.TAG_NAME, "tr")
        headers = header_row.find_elements(By.TAG_NAME, "th")
        
        # Map headers to their column indices
        header_map = {}
        for idx, header in enumerate(headers):
            header_text = extract_cell_text(header).strip()
            if header_text:
                header_map[header_text] = idx
                print(f"   Column {idx}: '{header_text}'")
        
        print(f"\n   Header mapping: {header_map}")
        
        # Get column indices for our target headers
        court_col = header_map.get("Court", 0)
        list_type_col = header_map.get("List Type", 1)
        round_col = header_map.get("Round", 2)
        sno_col = header_map.get("SNo.", 3)
        case_col = header_map.get("Case No.", 4)
        purpose_col = header_map.get("Purpose", 5)
        
        print(f"   Target columns -> Court:{court_col}, List:{list_type_col}, Round:{round_col}, SNo:{sno_col}, Case:{case_col}, Purpose:{purpose_col}")
        
        all_courts_data = []
        
        # Get tbody
        tbody = table.find_element(By.TAG_NAME, "tbody")
        rows = tbody.find_elements(By.TAG_NAME, "tr")
        
        print(f"\n   → Found {len(rows)} total rows")
        
        # Track current values for rowspan cells
        current_court = ""
        
        # Process rows (skip marquee rows)
        row_count = 0
        i = 0
        while i < len(rows):
            try:
                row = rows[i]
                cells = row.find_elements(By.TAG_NAME, "td")
                
                # Check if this is a marquee row (has colspan)
                if len(cells) == 1:
                    # This is a marquee row, skip it
                    i += 1
                    continue
                
                # Skip if not enough cells
                if len(cells) < 5:
                    i += 1
                    continue
                
                row_count += 1
                
                # Check if first cell has rowspan (new court)
                first_cell = cells[0]
                rowspan_attr = first_cell.get_attribute("rowspan")
                
                # Determine actual column positions based on whether Court cell exists
                if rowspan_attr and int(rowspan_attr) > 1:
                    # Row HAS Court cell - use direct mapping
                    current_court = extract_cell_text(cells[0]) if len(cells) > 0 else ""
                    list_type = extract_cell_text(cells[1]) if len(cells) > 1 else ""
                    round_val = extract_cell_text(cells[2]) if len(cells) > 2 else ""
                    sno = extract_cell_text(cells[3]) if len(cells) > 3 else ""
                    case_full = extract_cell_text(cells[4]) if len(cells) > 4 else ""
                    purpose = extract_cell_text(cells[5]) if len(cells) > 5 else ""
                else:
                    # Row DOESN'T have Court cell - shift indices left by 1
                    # Cells are: List Type | Round | SNo | Case No | Purpose
                    list_type = extract_cell_text(cells[0]) if len(cells) > 0 else ""
                    round_val = extract_cell_text(cells[1]) if len(cells) > 1 else ""
                    sno = extract_cell_text(cells[2]) if len(cells) > 2 else ""
                    case_full = extract_cell_text(cells[3]) if len(cells) > 3 else ""
                    purpose = extract_cell_text(cells[4]) if len(cells) > 4 else ""
                
                # Parse case details
                case_number, case_type, case_year = parse_case_details(case_full)
                
                # Create court record
                court_data = {
                    "Court": current_court,
                    "List Type": list_type,
                    "Round": round_val,
                    "SNo.": sno,
                    "Case No.": case_number,
                    "Case Type": case_type,
                    "Case Year": case_year,
                    "Full Case": case_full,
                    "Purpose": purpose,
                    "DateTime": scrape_time
                }
                
                all_courts_data.append(court_data)
                print(f"      ✓ Row {row_count}: Court='{current_court}' List='{list_type}' Case={case_number} ({case_type}/{case_year})")
                
                i += 1
                
            except Exception as e:
                print(f"      ✗ Error at row {i}: {str(e)}")
                import traceback
                traceback.print_exc()
                i += 1
                continue
        
        print(f"\n{'='*100}")
        print(f"EXTRACTION SUMMARY:")
        print(f"{'='*100}")
        print(f"   ✓ Total records extracted: {len(all_courts_data)}")
        print(f"   ✓ Timestamp: {scrape_time}")
        
        if all_courts_data:
            print(f"\n   Sample extracted data (first 5 rows):")
            sample_size = min(5, len(all_courts_data))
            for i, court in enumerate(all_courts_data[:sample_size], 1):
                print(f"      {i}. Court='{court['Court']}' | List='{court['List Type']}' | Round='{court['Round']}' | SNo='{court['SNo.']}' | Case={court['Case No.']} ({court['Case Type']}/{court['Case Year']})")
        
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
        df = df[["Court", "List Type", "Round", "SNo.", "Case No.", "Case Type", "Case Year", "Full Case", "Purpose", "DateTime"]]
        
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
        import traceback
        traceback.print_exc()
        return False


# ==================== MAIN EXECUTION ====================

def main():
    """Main execution"""
    print("=" * 100)
    print(" " * 20 + "CHHATTISGARH HIGH COURT DISPLAY BOARD SCRAPER")
    print(" " * 20 + "WITH EXCEL BACKUP + API INTEGRATION")
    print(" " * 20 + "API MAPPING: Purpose->caseNumber | Full Case->serialNumber")
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
        print(f"   Excel 'Purpose' (e.g., 'CRMP / 3087 / 2025') -> API 'caseNumber' (e.g., '3087')")
        print(f"   Excel 'Full Case' -> API 'serialNumber'")
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