"""
Andhra Pradesh High Court Display Board Scraper - FIXED VERSION
With Excel Backup + API Integration
Extracts ALL courts including Not in session
Handles dynamically loaded content via JavaScript
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
URL = "https://aphc.gov.in/Hcdbs/displayboard.jsp"
SCRAPE_INTERVAL = 30  # seconds
BASE_FOLDER = r"D:\CourtDisplayBoardScraper\displayboardexcel\andhra_pradesh_hc_excel"
BACKUP_CYCLE_INTERVAL = 60  # Create backup after every 60 cycles
SUB_BENCH_NO = "3"  # Sub-bench number for Andhra Pradesh HC
BENCH_NAME = "Amaravati"

# API Configuration
API_URL = "https://api.courtlivestream.com/api/display-boards/create"
API_TIMEOUT = 10  # seconds
ENABLE_API_POSTING = True  # Set to False to disable API posting
ENABLE_EXCEL_SAVING = True  # Set to False to disable Excel saving

# ==================== API FUNCTIONS ====================

def post_court_data_to_api(court_data):
    """Post a single court record to the API (NO TOKEN - as per requirement)"""
    try:
        # Extract date and time from DateTime field
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
        
        # Convert Item No to integer (serialNumber)
        serial_number = court_data.get("Item No", "")
        try:
            # Extract numeric part from Item No (e.g., "D-10" -> 10)
            if serial_number and isinstance(serial_number, str):
                numeric_part = re.search(r'\d+', serial_number)
                if numeric_part:
                    serial_number = int(numeric_part.group())
                else:
                    serial_number = 0
            else:
                serial_number = int(serial_number) if serial_number else 0
        except (ValueError, TypeError):
            serial_number = 0
        
        # For listNumber, we'll use a default value of 0
        list_number = 0
        
        # Extract ONLY the court number
        court_no_full = court_data.get("Court No", "")
        court_hall_number = court_no_full
        
        # Extract just the number part
        if court_no_full:
            match = re.match(r'^(\d+)', str(court_no_full))
            if match:
                court_hall_number = match.group(1)
        
        # Prepare API payload
        payload = {
            "benchName": court_data.get("Bench Name", ""),
            "courtHallNumber": court_hall_number,
            "caseNumber": "",  # AP HC doesn't provide case number on display board
            "serialNumber": serial_number,
            "date": date_str,
            "time": time_str,
            "passedOverCases": court_data.get("Kept Back Cases", ""),
            "listNumber": list_number
        }
        
        # Set headers (NO Authorization token)
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
        court_no = court_data.get("Court No", "N/A")
        item_no = court_data.get("Item No", "N/A")
        
        # Extract just the court number for display
        court_num = court_no
        if court_no and court_no != "N/A":
            match = re.match(r'^(\d+)', str(court_no))
            if match:
                court_num = match.group(1)
        
        print(f"   [{idx}/{total_courts}] Court {court_num} (Item: {item_no})...", end=" ")
        
        success, response = post_court_data_to_api(court_data)
        
        if success:
            successful_posts += 1
            print("✓")
        else:
            failed_posts += 1
            print(f"✗ ({response})")
            errors.append({"court": court_num, "item": item_no, "error": response})
    
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
    """Create date-based folder structure"""
    if not ENABLE_EXCEL_SAVING:
        return None
        
    current_date = datetime.now().strftime("%Y_%m_%d")
    date_folder = os.path.join(BASE_FOLDER, f"amaravati_{current_date}")
    
    if not os.path.exists(date_folder):
        os.makedirs(date_folder)
        print(f"✓ Created folder: {date_folder}")
    
    return date_folder


def get_date_folder():
    """Get today's date-based folder path"""
    current_date = datetime.now().strftime("%Y_%m_%d")
    date_folder = os.path.join(BASE_FOLDER, f"amaravati_{current_date}")
    return date_folder


def get_excel_path(folder):
    """Get full path for today's main Excel file"""
    if not folder:
        return None
    current_date = datetime.now().strftime("%Y_%m_%d")
    filename = f"amaravati_{current_date}.xlsx"
    excel_path = os.path.join(folder, filename)
    return excel_path


def get_timestamped_backup_path(folder):
    """Get full path for timestamped backup Excel file"""
    if not folder:
        return None
    current_timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M")
    filename = f"amaravati_bk_{current_timestamp}.xlsx"
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

def wait_for_table_to_populate(driver, max_wait=30):
    """
    Wait for the JavaScript to populate the table with actual data
    Returns True if data loaded, False if timeout
    """
    print("   → Waiting for JavaScript to populate table with data...")
    
    wait_count = 0
    while wait_count < max_wait:
        try:
            # Execute JavaScript to check if tbody has meaningful content
            has_data = driver.execute_script("""
                var tbody = document.getElementById('tbody');
                if (!tbody) return false;
                
                var rows = tbody.getElementsByTagName('tr');
                if (rows.length === 0) return false;
                
                // Check if first row has actual court data (not just empty structure)
                var firstRow = rows[0];
                var cells = firstRow.getElementsByTagName('td');
                
                if (cells.length === 0) return false;
                
                // Look for court number link in first cell
                var firstCell = cells[0];
                var hasCourtLink = firstCell.innerHTML.includes('getCourtCauseList');
                
                return hasCourtLink;
            """)
            
            if has_data:
                # Get row count for confirmation
                row_count = driver.execute_script("""
                    var tbody = document.getElementById('tbody');
                    return tbody ? tbody.getElementsByTagName('tr').length : 0;
                """)
                print(f"   ✓ Table populated with {row_count} rows after {wait_count} seconds")
                return True
                
        except Exception as e:
            print(f"   Debug: Error checking table: {str(e)}")
        
        time.sleep(1)
        wait_count += 1
        
        if wait_count % 5 == 0:
            print(f"   ⏳ Still waiting... ({wait_count}s elapsed)")
    
    print(f"   ✗ Timeout after {max_wait} seconds waiting for table data")
    return False


def scrape_display_board(driver):
    """
    Scrape courts from Andhra Pradesh HC display board
    Handles dynamically loaded JavaScript content - IMPROVED VERSION
    """
    try:
        print("   → Loading display board page...")
        driver.get(URL)
        
        # Wait for table element to exist
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "table1"))
        )
        print("   ✓ Table element found")
        
        # Wait for tbody to exist
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "tbody"))
        )
        print("   ✓ Tbody element found")
        
        # CRITICAL FIX: Wait for JavaScript to actually populate the table
        if not wait_for_table_to_populate(driver, max_wait=30):
            print("   ⚠ WARNING: Table may not be fully populated")
            # Continue anyway to see what we can extract
        
        # Additional small wait to ensure rendering is complete
        time.sleep(2)
        
        scrape_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        print("\n" + "="*100)
        print("ANALYZING PAGE STRUCTURE - EXTRACTING ALL COURTS...")
        print("="*100)
        
        # Get tbody element and check its content
        tbody = driver.find_element(By.ID, "tbody")
        
        # Debug: Print tbody HTML (first 500 chars)
        tbody_html = tbody.get_attribute('innerHTML')
        print(f"\n   DEBUG - Tbody HTML preview (first 500 chars):")
        print(f"   {tbody_html[:500]}...")
        print()
        
        rows = tbody.find_elements(By.TAG_NAME, "tr")
        
        print(f"   → Found {len(rows)} data rows in tbody")
        
        if len(rows) == 0:
            print("\n   ✗ ERROR: No rows found in tbody!")
            print("   This usually means:")
            print("      1. JavaScript hasn't loaded the data yet (needs more wait time)")
            print("      2. The website's AJAX endpoint is not responding")
            print("      3. The website structure has changed")
            return []
        
        all_courts_data = []
        
        print(f"\n{'─'*100}")
        print("EXTRACTING DATA FROM ROWS (ALL COURTS):")
        print(f"{'─'*100}")
        
        # Process each row
        for row_idx, row in enumerate(rows, 1):
            cells = row.find_elements(By.TAG_NAME, "td")
            
            if len(cells) < 4:
                print(f"   Row {row_idx}: Skipped (less than 4 cells: {len(cells)})")
                continue
            
            # Debug first row
            if row_idx == 1:
                print(f"\n   DEBUG - First row has {len(cells)} cells")
                for i, cell in enumerate(cells[:8]):  # Show first 8 cells
                    print(f"      Cell {i}: {extract_cell_text(cell)[:50]}")
                print()
            
            # Each row has 4 sets of: Court No, Coram/Status, Item No, Kept Back Cases
            # Some courts have "Not in session" with colspan=3
            idx = 0
            court_in_row = 0
            
            while idx < len(cells):
                try:
                    # Check if this cell has a court number link
                    cell_html = cells[idx].get_attribute('innerHTML')
                    has_court_link = 'getCourtCauseList' in cell_html
                    
                    if has_court_link:
                        court_in_row += 1
                        
                        # Extract court number
                        court_no = extract_cell_text(cells[idx])
                        
                        # Check if next cell has colspan (indicates "Not in session")
                        if idx + 1 < len(cells):
                            colspan = cells[idx + 1].get_attribute('colspan')
                            next_cell_text = extract_cell_text(cells[idx + 1])
                            
                            if colspan and int(colspan) >= 3:
                                # This is "Not in session" or "Session Started" - 2 cells total
                                coram = ""
                                item_no = next_cell_text  # Contains status message
                                kept_back = ""
                                
                                court_data = {
                                    "Bench Name": BENCH_NAME,
                                    "SubBenchNo": SUB_BENCH_NO,
                                    "Court No": court_no,
                                    "Coram": coram,
                                    "Item No": item_no,
                                    "Kept Back Cases": kept_back,
                                    "DateTime": scrape_time
                                }
                                all_courts_data.append(court_data)
                                
                                idx += 2  # Move past Court No + Status cells
                            else:
                                # Normal court with all 4 cells
                                coram = next_cell_text
                                item_no = ""
                                kept_back = ""
                                
                                if idx + 2 < len(cells):
                                    item_no = extract_cell_text(cells[idx + 2])
                                if idx + 3 < len(cells):
                                    kept_back = extract_cell_text(cells[idx + 3])
                                
                                court_data = {
                                    "Bench Name": BENCH_NAME,
                                    "SubBenchNo": SUB_BENCH_NO,
                                    "Court No": court_no,
                                    "Coram": coram,
                                    "Item No": item_no,
                                    "Kept Back Cases": kept_back,
                                    "DateTime": scrape_time
                                }
                                all_courts_data.append(court_data)
                                
                                idx += 4  # Move past all 4 cells
                        else:
                            idx += 1
                    else:
                        idx += 1
                        
                except Exception as e:
                    print(f"   Row {row_idx}, Cell {idx}: Error - {str(e)}")
                    idx += 1
                    continue
            
            if court_in_row > 0:
                print(f"   Row {row_idx}: Extracted {court_in_row} courts")
            else:
                print(f"   Row {row_idx}: No courts extracted (may be empty or malformed)")
        
        print(f"\n{'='*100}")
        print(f"EXTRACTION SUMMARY:")
        print(f"{'='*100}")
        print(f"   ✓ Total courts extracted: {len(all_courts_data)}")
        print(f"   ✓ Timestamp: {scrape_time}")
        
        # Show sample of extracted courts
        if all_courts_data:
            print(f"\n   Sample of extracted courts (first 5):")
            for i, court in enumerate(all_courts_data[:5], 1):
                status = court['Item No'] if court['Coram'] == '' else f"{court['Coram']} - {court['Item No']}"
                print(f"      {i}. Court {court['Court No']}: {status}")
        else:
            print(f"\n   ⚠ WARNING: No courts were extracted!")
            print(f"   Possible issues:")
            print(f"      - Page structure has changed")
            print(f"      - JavaScript didn't execute properly")
            print(f"      - Network/AJAX request failed")
        
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
        df = df[["Bench Name", "SubBenchNo", "Court No", "Coram", "Item No", "Kept Back Cases", "DateTime"]]
        
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
    print(" " * 20 + "ANDHRA PRADESH HIGH COURT DISPLAY BOARD SCRAPER - FIXED")
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
        print(f"Authentication: NO TOKEN (as per requirement)")
    print(f"SubBench Number: {SUB_BENCH_NO}")
    print(f"Bench Name: {BENCH_NAME}")
    print("\nFIXES APPLIED:")
    print("   ✓ Enhanced wait for JavaScript-loaded content")
    print("   ✓ Added table population detection")
    print("   ✓ Improved debugging output")
    print("   ✓ Better error handling for dynamic content")
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
                    print(f"   Old folder: {os.path.basename(date_folder)}")
                    print(f"   New folder: {os.path.basename(current_date_folder)}")
                    print(f"{'='*100}\n")
                    
                    date_folder = create_folder()
                    excel_path = get_excel_path(date_folder)
                    first_cycle = True
                    last_backup_cycle = 0
                    cycle_count = 1
                    
                    print(f"✓ New main file: {os.path.basename(excel_path)}")
            
            print(f"\n{'='*100}")
            print(f"CYCLE {cycle_count} - {current_time}")
            if ENABLE_EXCEL_SAVING and excel_path:
                print(f"Folder: {os.path.basename(date_folder)}")
                print(f"Main Excel: {os.path.basename(excel_path)}")
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
                print(f"   Extracted: {len(courts_data)} courts from Amaravati Bench")
                
                if ENABLE_EXCEL_SAVING:
                    status = "SUCCESS" if excel_success else "FAILED"
                    print(f"   Excel Save: {status}")
                
                if ENABLE_API_POSTING and api_result:
                    print(f"   API Posting: {api_result['successful']}/{api_result['total']} successful")
                
                print(f"{'='*100}")
                
                if excel_success:
                    first_cycle = False
                    
                    if ENABLE_EXCEL_SAVING and cycle_count - last_backup_cycle >= BACKUP_CYCLE_INTERVAL:
                        print(f"\n{'─'*100}")
                        print(f"⏰ BACKUP TIME - {BACKUP_CYCLE_INTERVAL} cycles completed")
                        print(f"   Creating timestamped backup from main Excel file")
                        print(f"{'─'*100}")
                        
                        backup_success = create_backup_from_main_excel(excel_path, date_folder)
                        
                        if backup_success:
                            last_backup_cycle = cycle_count
                            print(f"   ✓ Backup created successfully")
            else:
                print(f"\n   ✗ No data scraped in cycle {cycle_count}")
                print(f"   ℹ Possible reasons:")
                print(f"      - JavaScript content not loading")
                print(f"      - AJAX endpoint not responding")
                print(f"      - Network connectivity issues")
                print(f"      - Website structure changed")
            
            next_time = datetime.fromtimestamp(time.time() + SCRAPE_INTERVAL).strftime('%Y-%m-%d %H:%M:%S')
            
            print(f"\n{'─'*100}")
            print(f"⏳ Waiting {SCRAPE_INTERVAL} seconds")
            print(f"   Next cycle: {next_time}")
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

