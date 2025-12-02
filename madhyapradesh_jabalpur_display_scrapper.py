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

# ==================== CONFIGURATION ====================
URL = "https://mphc.gov.in/online-display-board"
SCRAPE_INTERVAL = 30  # seconds
BASE_FOLDER = r"D:\CourtDisplayBoardScraper\displayboardexcel\madhyapradesh_hc_excels\jabalpur_bench"
BENCH_VALUE = "01"  # Jabalpur bench value
BACKUP_CYCLE_INTERVAL = 60  # Create backup after every 3 cycles
SUB_BENCH_NO = "26"  # Sub-bench number for Jabalpur
BENCH_NAME = "jabalpur"
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
    Format: D:\CourtDisplayBoardScraper\displayboardexcel\madhyapradesh_hc_excels\jabalpur_bench\jabalpur_YYYY_MM_DD\
    """
    current_date = datetime.now().strftime("%Y_%m_%d")
    date_folder = os.path.join(BASE_FOLDER, f"jabalpur_{current_date}")
    
    if not os.path.exists(date_folder):
        os.makedirs(date_folder)
        print(f"✓ Created folder: {date_folder}")
    
    return date_folder


def get_date_folder():
    """Get today's date-based folder path"""
    current_date = datetime.now().strftime("%Y_%m_%d")
    date_folder = os.path.join(BASE_FOLDER, f"jabalpur_{current_date}")
    return date_folder


def get_excel_path(folder):
    """
    Get full path for today's main Excel file
    Format: jabalpur_YYYY_MM_DD.xlsx
    """
    current_date = datetime.now().strftime("%Y_%m_%d")
    filename = f"jabalpur_{current_date}.xlsx"
    excel_path = os.path.join(folder, filename)
    return excel_path


def get_timestamped_backup_path(folder):
    """
    Get full path for timestamped backup Excel file
    Format: jabalpur_bk_YYYY_MM_DD_HH_MM.xlsx
    """
    current_timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M")
    filename = f"jabalpur_bk_{current_timestamp}.xlsx"
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


def extract_court_number(cell):
    """
    Extract ONLY the court number from the first cell
    The cell contains: <strong><font>2</font></strong>
    We need to extract only the number
    """
    try:
        # Try to find the <strong><font> tag that contains the court number
        strong_tag = cell.find_element(By.TAG_NAME, "strong")
        font_tag = strong_tag.find_element(By.TAG_NAME, "font")
        court_no = font_tag.text.strip()
        return court_no
    except:
        # Fallback: try to get text content
        try:
            text = cell.text.strip()
            # Extract only the first number (court number)
            lines = text.split('\n')
            if lines:
                return lines[0].strip()
            return text
        except:
            return ""


# ==================== SCRAPING FUNCTIONS ====================

def select_bench(driver):
    """Select Jabalpur bench from dropdown"""
    try:
        print("   → Selecting Jabalpur bench from dropdown...")
        
        # Wait for dropdown to be present
        dropdown = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "my_city"))
        )
        
        # Select Jabalpur (value = "01")
        select = Select(dropdown)
        select.select_by_value(BENCH_VALUE)
        
        print("   ✓ Jabalpur bench selected")
        
        # Wait for page to reload with new data
        time.sleep(5)
        
        return True
    except Exception as e:
        print(f"   ✗ Error selecting bench: {str(e)}")
        return False


def scrape_display_board(driver):
    """
    Scrape courts from Madhya Pradesh High Court - JABALPUR BENCH
    Columns: Court No. | Sr. No. | Case. No. | Petitioner | Respondent | Court Message
    """
    try:
        print("   → Loading display board page...")
        driver.get(URL)
        
        # Wait for page to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "my_city"))
        )
        
        # Select Jabalpur bench
        if not select_bench(driver):
            return []
        
        # Wait for table to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "board_id"))
        )
        time.sleep(3)
        
        # Get current timestamp
        scrape_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        print("\n" + "="*100)
        print("EXTRACTING JABALPUR BENCH COURTS...")
        print("="*100)
        
        all_courts_data = []
        
        # Find the data table
        tables = driver.find_elements(By.CLASS_NAME, "board_id")
        print(f"   → Found {len(tables)} table(s) on the page")
        
        if len(tables) > 0:
            table = tables[0]  # Get the first table
            
            rows = table.find_elements(By.TAG_NAME, "tr")
            print(f"   → Found {len(rows)} total rows in table")
            
            # Check headers
            if len(rows) > 0:
                print(f"\n   HEADER ROW:")
                header_row = rows[0]
                header_cells = header_row.find_elements(By.TAG_NAME, "th")
                
                for idx, cell in enumerate(header_cells):
                    header_text = extract_cell_text(cell)
                    print(f"      Header[{idx}]: '{header_text}'")
            
            print(f"\n   EXTRACTING DATA FROM ALL ROWS:")
            
            # Process each data row (skip header)
            for row_idx, row in enumerate(rows[1:], 1):
                try:
                    # Check if this is a data row (class="record")
                    if "record" not in row.get_attribute("class"):
                        continue
                    
                    cells = row.find_elements(By.TAG_NAME, "td")
                    
                    if len(cells) >= 7:  # Ensure we have all columns
                        # Extract data from each column
                        # Column 0: Court No. - EXTRACT ONLY THE NUMBER
                        court_no = extract_court_number(cells[0])
                        
                        # Column 3: Sr. No.
                        sr_no = extract_cell_text(cells[3])
                        
                        # Column 4: Case No.
                        case_no = extract_cell_text(cells[4])
                        
                        # Column 5: Petitioner
                        petitioner = extract_cell_text(cells[5])
                        
                        # Column 6: Respondent
                        respondent = extract_cell_text(cells[6])
                        
                        # Column 7: Court Message
                        court_message = extract_cell_text(cells[7]) if len(cells) > 7 else ""
                        
                        print(f"\n      ROW {row_idx} (Court {court_no}):")
                        print(f"         Court No: '{court_no}'")
                        print(f"         Sr. No: '{sr_no}'")
                        print(f"         Case No: '{case_no}'")
                        print(f"         Petitioner: '{petitioner}'")
                        print(f"         Respondent: '{respondent}'")
                        print(f"         Court Message: '{court_message}'")
                        
                        # Create court data dictionary
                        court_data = {
                            "Bench Name": BENCH_NAME,
                            "SubBenchNo": SUB_BENCH_NO,
                            "Court No": court_no if court_no else "",
                            "Sr. No": sr_no if sr_no else "",
                            "Case No": case_no if case_no else "",
                            "Petitioner": petitioner if petitioner else "",
                            "Respondent": respondent if respondent else "",
                            "Court Message": court_message if court_message else "",
                            "DateTime": scrape_time
                        }
                        
                        all_courts_data.append(court_data)
                        print(f"         ✓ EXTRACTED")
                    else:
                        print(f"      Row {row_idx}: Warning - only {len(cells)} cells found")
                        
                except Exception as e:
                    print(f"      ✗ Error processing row {row_idx}: {str(e)}")
                    continue
        else:
            print(f"\n   ✗ ERROR: Could not find data table")
        
        print(f"\n{'='*100}")
        print(f"EXTRACTION SUMMARY:")
        print(f"{'='*100}")
        print(f"   ✓ Total courts extracted: {len(all_courts_data)}")
        print(f"   ✓ Timestamp: {scrape_time}")
        
        if all_courts_data:
            print(f"\n   All extracted courts from Jabalpur Bench:")
            for i, court in enumerate(all_courts_data, 1):
                print(f"      {i}. Court {court['Court No']} | Sr {court['Sr. No']} | {court['Case No']}")
        
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
        df = df[["Bench Name","SubBenchNo", "Court No", "Sr. No", "Case No", "Petitioner", "Respondent", "Court Message", "DateTime"]]
        
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
    print(" " * 20 + "MADHYA PRADESH HIGH COURT - JABALPUR BENCH DISPLAY BOARD SCRAPER")
    print(" " * 30 + "WITH TIMESTAMPED BACKUP FILES EVERY 3 CYCLES")
    print("=" * 100)
    print(f"URL: {URL}")
    print(f"Scrape Interval: {SCRAPE_INTERVAL} seconds")
    print(f"Base Location: {BASE_FOLDER}")
    print(f"Backup Interval: Every {BACKUP_CYCLE_INTERVAL} cycles")
    print(f"Target: Jabalpur Bench ONLY")
    print(f"Folder Structure:")
    print(f"   D:\\CourtDisplayBoardScraper\\displayboardexcel\\madhyapradesh_hc_excels\\")
    print(f"   └── jabalpur_bench\\")
    print(f"       └── jabalpur_YYYY_MM_DD\\")
    print(f"           ├── jabalpur_YYYY_MM_DD.xlsx (main file)")
    print(f"           ├── jabalpur_bk_YYYY_MM_DD_HH_MM.xlsx (backup after 3 cycles)")
    print(f"           ├── jabalpur_bk_YYYY_MM_DD_HH_MM.xlsx (backup after 6 cycles)")
    print(f"           └── jabalpur_bk_YYYY_MM_DD_HH_MM.xlsx (backup after 9 cycles)")
    print(f"           etc...")
    print(f"SubBench Number: {SUB_BENCH_NO} (applied to all records)")
    print(f"Bench Name:{BENCH_NAME} (applied to all records)")
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
                    print(f"   Extracted {len(courts_data)} courts from Jabalpur Bench")
                    print(f"{'='*100}")
                    first_cycle = False
                    
                    # Check if backup is needed (every 3 cycles)
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