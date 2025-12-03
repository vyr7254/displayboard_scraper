"""
Patna High Court Display Board Scraper
Extracts court data from Patna HC display board and saves to Excel
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

# ==================== CONFIGURATION ====================
URL = "https://patnahighcourt.gov.in/online_display_board"
SCRAPE_INTERVAL = 30  # seconds
BASE_FOLDER = r"D:\CourtDisplayBoardScraper\displayboardexcel\patna_hc_excel\patna_bench"
BACKUP_CYCLE_INTERVAL = 60  # Create backup after every 60 cycles
SUB_BENCH_NO = "34"  # Sub-bench number for Patna
BENCH_NAME = "patna"

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
    current_date = datetime.now().strftime("%Y_%m_%d")
    filename = f"patna_{current_date}.xlsx"
    excel_path = os.path.join(folder, filename)
    return excel_path


def get_timestamped_backup_path(folder):
    """
    Get full path for timestamped backup Excel file
    Format: patna_bk_YYYY_MM_DD_HH_MM.xlsx
    """
    current_timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M")
    filename = f"patna_bk_{current_timestamp}.xlsx"
    backup_path = os.path.join(folder, filename)
    return backup_path


def create_backup_from_main_excel(main_excel_path, folder):
    """
    Create a timestamped backup file by copying ALL data from the main Excel file
    This ensures we have a complete snapshot at the time of backup
    """
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
    print(" " * 30 + "WITH TIMESTAMPED BACKUP FILES EVERY 60 CYCLES")
    print("=" * 100)
    print(f"URL: {URL}")
    print(f"Scrape Interval: {SCRAPE_INTERVAL} seconds")
    print(f"Base Location: {BASE_FOLDER}")
    print(f"Backup Interval: Every {BACKUP_CYCLE_INTERVAL} cycles")
    print(f"Folder Structure:")
    print(f"   D:\\CourtDisplayBoardScraper\\displayboardexcel\\")
    print(f"   └── patna_hc_excel\\")
    print(f"       └── patna_bench\\")
    print(f"           └── patna_YYYY_MM_DD\\")
    print(f"               ├── patna_YYYY_MM_DD.xlsx (main file)")
    print(f"               ├── patna_bk_YYYY_MM_DD_HH_MM.xlsx (backup after 60 cycles)")
    print(f"               ├── patna_bk_YYYY_MM_DD_HH_MM.xlsx (backup after 120 cycles)")
    print(f"               └── patna_bk_YYYY_MM_DD_HH_MM.xlsx (backup after 180 cycles)")
    print(f"               etc...")
    print(f"SubBench Number: {SUB_BENCH_NO} (applied to all records)")
    print(f"Bench Name: {BENCH_NAME} (applied to all records)")
    print("=" * 100)
    
    # Get today's folder and file paths
    date_folder = create_folder()
    excel_path = get_excel_path(date_folder)
    
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
            
            # Check if date has changed (new day started)
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
            print(f"Folder: {os.path.basename(date_folder)}")
            print(f"Main Excel: {os.path.basename(excel_path)}")
            print(f"{'='*100}")
            
            courts_data = scrape_display_board(driver)
            
            if courts_data:
                # Save to main Excel file
                success = save_to_excel(courts_data, excel_path, open_file=first_cycle)
                
                if success:
                    print(f"\n{'='*100}")
                    print(f"✓✓✓ CYCLE {cycle_count} COMPLETED SUCCESSFULLY ✓✓✓")
                    print(f"   Extracted {len(courts_data)} courts from Patna Bench")
                    print(f"{'='*100}")
                    first_cycle = False
                    
                    # Check if backup is needed (every 60 cycles)
                    if cycle_count - last_backup_cycle >= BACKUP_CYCLE_INTERVAL:
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
                    print(f"\n   ⚠ Save failed in cycle {cycle_count}")
            else:
                print(f"\n   ✗ No data scraped in cycle {cycle_count}")
            
            next_time = datetime.fromtimestamp(time.time() + SCRAPE_INTERVAL).strftime('%Y-%m-%d %H:%M:%S')
            cycles_until_backup = BACKUP_CYCLE_INTERVAL - (cycle_count - last_backup_cycle)
            
            print(f"\n{'─'*100}")
            print(f"⏳ Waiting {SCRAPE_INTERVAL} seconds")
            print(f"   Next cycle: {next_time}")
            print(f"   Next backup in: {cycles_until_backup} cycle(s)")
            print(f"{'─'*100}")
            time.sleep(SCRAPE_INTERVAL)
    
    except KeyboardInterrupt:
        print("\n" + "=" * 100)
        print("⚠ Script stopped by user")
        print(f"Total cycles completed: {cycle_count}")
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