"""
Orissa High Court Display Board Scraper
Extracts court data from Orissa HC display board and saves to Excel
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
BASE_FOLDER = r"D:\CourtDisplayBoardScraper\displayboardexcel\orissa_hc_excels"
EXCEL_FILE = "OrissaHC_DisplayBoard_Data.xlsx"

# ==================== SETUP FUNCTIONS ====================

def setup_driver():
    """
    Initialize Chrome driver with VISIBLE browser
    """
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager
    
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36")
    
    # Use webdriver-manager for automatic driver management
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.implicitly_wait(10)
    return driver


def create_folder():
    """
    Create folder structure if it doesn't exist
    """
    if not os.path.exists(BASE_FOLDER):
        os.makedirs(BASE_FOLDER)
        print(f"✓ Created folder: {BASE_FOLDER}")
    
    excel_path = os.path.join(BASE_FOLDER, EXCEL_FILE)
    return excel_path


def open_excel_file(file_path):
    """
    Open Excel file automatically after first save
    """
    try:
        if platform.system() == 'Windows':
            os.startfile(file_path)
            print(f"   ✓ Excel file opened: {file_path}")
    except Exception as e:
        print(f"   ⚠ Could not auto-open Excel: {str(e)}")


def extract_cell_text(cell):
    """
    Extract visible text from cell, handling nested HTML elements
    """
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
            for i, court in enumerate(all_courts_data[:5], 1):
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
    """
    Save scraped data to Excel
    """
    try:
        if not data:
            print("   ⚠ No data to save")
            return False
        
        # Convert to DataFrame
        df = pd.DataFrame(data)
        
        # Ensure column order
        df = df[["Court No", "Judge Name", "Sl.No", "Case Number", "Full Case Details", "DateTime"]]
        
        # Check if file exists
        if os.path.exists(file_path):
            # Read existing data
            existing_df = pd.read_excel(file_path, engine='openpyxl')
            
            # Concatenate
            combined_df = pd.concat([existing_df, df], ignore_index=True)
            
            # Write back
            combined_df.to_excel(file_path, index=False, sheet_name='Sheet1', engine='openpyxl')
            
            print(f"\n   ✓ Data appended to Excel")
            print(f"   ✓ Added {len(df)} courts (Total: {len(combined_df)} rows)")
            print(f"   ✓ File: {file_path}")
        else:
            # Create new file
            df.to_excel(file_path, index=False, sheet_name='Sheet1', engine='openpyxl')
            print(f"\n   ✓ New Excel file created")
            print(f"   ✓ Initial data: {len(df)} courts")
            print(f"   ✓ File: {file_path}")
        
        # Open Excel on first save
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
    """
    Main execution
    """
    print("=" * 100)
    print(" " * 30 + "ORISSA HIGH COURT DISPLAY BOARD SCRAPER")
    print("=" * 100)
    print(f"URL: {URL}")
    print(f"Scrape Interval: {SCRAPE_INTERVAL} seconds")
    print(f"Save Location: {BASE_FOLDER}")
    print("=" * 100)
    
    # Create folder
    excel_path = create_folder()
    print(f"✓ Excel file path: {excel_path}")
    print("=" * 100)
    
    # Initialize driver
    print("\nInitializing Chrome driver...")
    driver = setup_driver()
    print("✓ Browser opened")
    print("=" * 100)
    
    cycle_count = 0
    first_cycle = True
    
    try:
        while True:
            cycle_count += 1
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            print(f"\n{'='*100}")
            print(f"CYCLE {cycle_count} - {current_time}")
            print(f"{'='*100}")
            
            # Scrape
            courts_data = scrape_display_board(driver)
            
            # Save to Excel
            if courts_data:
                success = save_to_excel(courts_data, excel_path, open_file=first_cycle)
                
                if success:
                    print(f"\n{'='*100}")
                    print(f"✓✓✓ CYCLE {cycle_count} COMPLETED SUCCESSFULLY ✓✓✓")
                    print(f"{'='*100}")
                    first_cycle = False
                else:
                    print(f"\n   ⚠ Save failed in cycle {cycle_count}")
            else:
                print(f"\n   ✗ No data scraped in cycle {cycle_count}")
            
            # Wait
            next_time = datetime.fromtimestamp(time.time() + SCRAPE_INTERVAL).strftime('%Y-%m-%d %H:%M:%S')
            print(f"\n{'─'*100}")
            print(f"⏳ Waiting {SCRAPE_INTERVAL} seconds")
            print(f"   Next cycle: {next_time}")
            print(f"{'─'*100}")
            time.sleep(SCRAPE_INTERVAL)
    
    except KeyboardInterrupt:
        print("\n" + "=" * 100)
        print("⚠ Script stopped by user")
        print(f"Total cycles: {cycle_count}")
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