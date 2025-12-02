"""
Orissa High Court Display Board Scraper
Extracts court data from Orissa HC display board and saves to Excel
WITH TIMESTAMPED BACKUP FILES EVERY 60 CYCLES
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
URL = "http://www.ohcdb.in/"
SCRAPE_INTERVAL = 30  # seconds
BASE_FOLDER = r"D:\CourtDisplayBoardScraper\displayboardexcel\orissa_hc_excel"
BACKUP_CYCLE_INTERVAL = 60  # Create backup after every 60 cycles
BENCH_NAME = "Orissa"
SUB_BENCH_NO = "33"

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
    """
    Create date-based folder structure
    Format: D:\CourtDisplayBoardScraper\displayboardexcel\orissa_hc_excel\orissa_hc_YYYY_MM_DD\
    """
    current_date = datetime.now().strftime("%Y_%m_%d")
    date_folder = os.path.join(BASE_FOLDER, f"orissa_hc_{current_date}")
    
    if not os.path.exists(date_folder):
        os.makedirs(date_folder)
        print(f"✓ Created folder: {date_folder}")
    
    return date_folder


def get_date_folder():
    """Get today's date-based folder path"""
    current_date = datetime.now().strftime("%Y_%m_%d")
    date_folder = os.path.join(BASE_FOLDER, f"orissa_hc_{current_date}")
    return date_folder


def get_excel_path(folder):
    """
    Get full path for today's main Excel file
    Format: orissa_hc_YYYY_MM_DD.xlsx
    """
    current_date = datetime.now().strftime("%Y_%m_%d")
    filename = f"orissa_hc_{current_date}.xlsx"
    excel_path = os.path.join(folder, filename)
    return excel_path


def get_timestamped_backup_path(folder):
    """
    Get full path for timestamped backup Excel file
    Format: orissa_hc_bk_YYYY_MM_DD_HH_MM.xlsx
    """
    current_timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M")
    filename = f"orissa_hc_bk_{current_timestamp}.xlsx"
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


def extract_slno_and_case(case_details):
    """
    Extract Sl.No and Case Number from case details
    Format examples:
    - "WKL : 4. WP(C) 19033/2023" -> Sl.No: 4, Case: WP(C) 19033/2023
    - "SUPL : 6. WP(C) 31878/2025" -> Sl.No: 6, Case: WP(C) 31878/2025
    - "Not in Session" -> Sl.No: "", Case: Not in Session
    """
    try:
        if not case_details or case_details.lower() == "not in session":
            return "", case_details
        
        # Pattern: "WORD : NUMBER. CASE_DETAILS"
        # Examples: "WKL : 4. WP(C) 19033/2023", "SUPL : 29. RSA 574/2025"
        match = re.search(r':\s*(\d+)\.\s*(.+)', case_details)
        if match:
            sl_no = match.group(1)  # Extract the number after colon
            case_no = match.group(2).strip()  # Extract everything after the dot
            return sl_no, case_no
        
        # If pattern doesn't match, return empty sl_no and full case_details
        return "", case_details
    except:
        return "", case_details


# ==================== SCRAPING FUNCTIONS ====================

def scrape_display_board(driver):
    """
    Scrape courts from Orissa High Court display board
    Layout: 3 tables side by side, each with pattern:
    - Header row with "Court No." and "Type : Case No"
    - Judge name row (colspan=2)
    - Court data row (Court No | Case details)
    """
    try:
        print("   → Loading display board page...")
        driver.get(URL)
        
        # Wait for tables to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, "table"))
        )
        time.sleep(5)  # Extra wait for dynamic content
        
        # Get current timestamp
        scrape_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        print("\n" + "="*100)
        print("ANALYZING PAGE STRUCTURE - EXTRACTING ALL COURTS...")
        print("="*100)
        
        all_courts_data = []
        
        # Find all tables with border="1"
        all_tables = driver.find_elements(By.TAG_NAME, "table")
        court_tables = [t for t in all_tables if t.get_attribute("border") == "1"]
        
        print(f"   → Found {len(court_tables)} court tables")
        
        # Process each table
        for table_idx, table in enumerate(court_tables, 1):
            print(f"\n{'─'*100}")
            print(f"PROCESSING TABLE {table_idx}:")
            print(f"{'─'*100}")
            
            rows = table.find_elements(By.TAG_NAME, "tr")
            print(f"   → {len(rows)} rows in table {table_idx}")
            
            # Skip header row (first row)
            i = 1
            while i < len(rows):
                try:
                    # Pattern: Judge row (colspan=2), then Court data row
                    judge_row = rows[i]
                    
                    # Check if this is a judge row (has colspan=2)
                    judge_cells = judge_row.find_elements(By.TAG_NAME, "td")
                    
                    if len(judge_cells) == 1:  # Judge row (colspan=2)
                        judge_name = extract_cell_text(judge_cells[0])
                        
                        # Next row should be the court data
                        if i + 1 < len(rows):
                            court_row = rows[i + 1]
                            court_cells = court_row.find_elements(By.TAG_NAME, "td")
                            
                            if len(court_cells) >= 2:
                                court_no = extract_cell_text(court_cells[0])
                                case_details = extract_cell_text(court_cells[1])
                                
                                # Extract Sl.No and Case Number from case details
                                sl_no, case_no = extract_slno_and_case(case_details)
                                
                                # Only add if court number exists
                                if court_no:
                                    court_data = {
                                        "Bench Name": BENCH_NAME,
                                        "Sub Bench No": SUB_BENCH_NO,
                                        "Court No": court_no,
                                        "Judge Name": judge_name,
                                        "Sl.No": sl_no,
                                        "Case Number": case_no,
                                        "Full Case Details": case_details,
                                        "DateTime": scrape_time
                                    }
                                    
                                    all_courts_data.append(court_data)
                                    print(f"      ✓ Court {court_no}: Sl.No {sl_no} | {case_no[:40]}...")
                            
                            i += 2  # Move past judge and court rows
                        else:
                            i += 1
                    else:
                        i += 1
                        
                except Exception as e:
                    print(f"      ✗ Error at row {i}: {str(e)}")
                    i += 1
                    continue
        
        print(f"\n{'='*100}")
        print(f"EXTRACTION SUMMARY:")
        print(f"{'='*100}")
        print(f"   ✓ Total courts extracted: {len(all_courts_data)}")
        print(f"   ✓ Timestamp: {scrape_time}")
        
        if all_courts_data:
            print(f"\n   Sample extracted data:")
            sample_size = min(3, len(all_courts_data))
            for i, court in enumerate(all_courts_data[:sample_size], 1):
                print(f"      {i}. Court {court['Court No']} | Sl.No {court['Sl.No']} | {court['Case Number'][:40]}")
        
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
        df = df[["Sub Bench No", "Bench Name", "Court No", "Judge Name", "Sl.No", "Case Number", "Full Case Details", "DateTime"]]
        
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
    print(" " * 25 + "ORISSA HIGH COURT DISPLAY BOARD SCRAPER")
    print(" " * 30 + "WITH TIMESTAMPED BACKUP FILES EVERY 60 CYCLES")
    print("=" * 100)
    print(f"URL: {URL}")
    print(f"Scrape Interval: {SCRAPE_INTERVAL} seconds")
    print(f"Base Location: {BASE_FOLDER}")
    print(f"Backup Interval: Every {BACKUP_CYCLE_INTERVAL} cycles")
    print(f"Folder Structure:")
    print(f"   D:\\CourtDisplayBoardScraper\\displayboardexcel\\")
    print(f"   └── orissa_hc_excel\\")
    print(f"       └── orissa_hc_YYYY_MM_DD\\")
    print(f"           ├── orissa_hc_YYYY_MM_DD.xlsx (main file)")
    print(f"           ├── orissa_hc_bk_YYYY_MM_DD_HH_MM.xlsx (backup after 60 cycles)")
    print(f"           ├── orissa_hc_bk_YYYY_MM_DD_HH_MM.xlsx (backup after 120 cycles)")
    print(f"           └── orissa_hc_bk_YYYY_MM_DD_HH_MM.xlsx (backup after 180 cycles)")
    print(f"           etc...")
    print(f"Bench Name: {BENCH_NAME} (applied to all records)")
    print(f"Sub Bench No: {SUB_BENCH_NO} (applied to all records)")
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
                    print(f"   Extracted {len(courts_data)} courts from Orissa HC")
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