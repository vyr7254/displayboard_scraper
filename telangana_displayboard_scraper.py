"""
Telangana High Court Display Board Scraper
Modified to extract ALL courts including NS and COURT SESSION ENDED
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
import subprocess
import platform
from bs4 import BeautifulSoup

# ==================== CONFIGURATION ====================
URL = "https://displayboard.tshc.gov.in/hcdbs/displayall"
SCRAPE_INTERVAL = 30  # seconds
BASE_FOLDER = r"D:\CourtDisplayBoardScraper\displayboardexcel\telangana_hc_excels"
EXCEL_FILE = "TSHC_DisplayBoard_Data_2December.xlsx"

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
        # Get the inner HTML
        html_content = cell.get_attribute('innerHTML')
        
        # Parse with BeautifulSoup to extract text properly
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # Get all text, strip whitespace
        text = soup.get_text(separator=' ', strip=True)
        
        # Clean up extra whitespace
        text = re.sub(r'\s+', ' ', text).strip()
        
        return text
    except:
        return ""


def get_court_number_from_link(cell):
    """
    Extract court number from the href link in the cell
    """
    try:
        html_content = cell.get_attribute('innerHTML')
        # Extract court number from URL like "court-view?court=1"
        match = re.search(r'court=(\d+)', html_content)
        if match:
            return match.group(1)
    except:
        pass
    return None


# ==================== SCRAPING FUNCTIONS ====================

def scrape_display_board(driver):
    """
    Scrape courts from display board - EXTRACT ALL COURTS including NS and COURT SESSION ENDED
    """
    try:
        print("   → Loading display board page...")
        driver.get(URL)
        
        # Wait for page to load
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, "table"))
        )
        time.sleep(5)  # Extra wait for dynamic content
        
        # Get current timestamp
        scrape_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        print("\n" + "="*90)
        print("ANALYZING PAGE STRUCTURE - EXTRACTING ALL COURTS...")
        print("="*90)
        
        # Find all tables
        tables = driver.find_elements(By.TAG_NAME, "table")
        print(f"   → Found {len(tables)} table(s) on the page")
        
        # Use the first table
        table = tables[0]
        rows = table.find_elements(By.TAG_NAME, "tr")
        print(f"   → Found {len(rows)} total rows")
        
        # Check headers
        if len(rows) > 0:
            print(f"\n{'─'*90}")
            print("HEADER ROW:")
            print(f"{'─'*90}")
            header_cells = rows[0].find_elements(By.TAG_NAME, "th")
            if not header_cells:
                header_cells = rows[0].find_elements(By.TAG_NAME, "td")
            
            for idx, cell in enumerate(header_cells):
                header_text = extract_cell_text(cell)
                print(f"   Header[{idx}]: '{header_text}'")
        
        # Extract court data
        all_courts_data = []
        
        print(f"\n{'─'*90}")
        print("EXTRACTING DATA FROM ROWS (ALL COURTS):")
        print(f"{'─'*90}")
        
        for row_idx, row in enumerate(rows[1:], 1):
            print(f"\n   ROW {row_idx}:")
            
            cells = row.find_elements(By.TAG_NAME, "td")
            print(f"      → Found {len(cells)} cells")
            
            # Debug: Show extracted text for each cell
            cell_texts = []
            for cell_idx, cell in enumerate(cells):
                text = extract_cell_text(cell)
                court_num = get_court_number_from_link(cell)
                cell_texts.append(text)
                print(f"         Cell[{cell_idx}]: '{text}' {f'(Court #{court_num})' if court_num else ''}")
            
            # Now extract courts - each court has 4 columns
            # Pattern: Court No (link) | Running Item | Case No | Passed Over
            if len(cells) >= 4:
                print(f"\n      → Extracting courts from this row...")
                
                # Process cells in groups of 4
                court_in_row = 0
                idx = 0
                
                while idx < len(cells):
                    # Check if this is a court group by looking for court number in first cell
                    court_num = get_court_number_from_link(cells[idx])
                    
                    if court_num:  # This is a court number cell
                        court_in_row += 1
                        
                        # Extract court number
                        court_no = extract_cell_text(cells[idx])
                        
                        # Check if next cell contains "Court Session Ended" (might be merged/span columns)
                        running_item = ""
                        case_no = ""
                        passed_over = ""
                        
                        if idx + 1 < len(cells):
                            running_item = extract_cell_text(cells[idx + 1])
                            
                            # Check if "Court Session Ended" is in this cell
                            if "Court Session Ended" in running_item:
                                # Court Session Ended - may span multiple columns
                                # Set it in Running Item and leave others empty
                                case_no = ""
                                passed_over = ""
                                cells_consumed = 2  # Court No + merged cell
                                
                                # Check if there are actually separate cells or it's merged
                                # Try to peek at next cells to see if they're part of this court
                                if idx + 2 < len(cells):
                                    next_cell_text = extract_cell_text(cells[idx + 2])
                                    # If next cell has a court number link, current court only has 2 cells
                                    if not get_court_number_from_link(cells[idx + 2]):
                                        # Next cell might be case_no or continuation
                                        if idx + 3 < len(cells):
                                            case_no = next_cell_text
                                            passed_over = extract_cell_text(cells[idx + 3])
                                            cells_consumed = 4
                                        else:
                                            cells_consumed = 2
                            else:
                                # Normal court with separate columns
                                if idx + 2 < len(cells):
                                    case_no = extract_cell_text(cells[idx + 2])
                                if idx + 3 < len(cells):
                                    passed_over = extract_cell_text(cells[idx + 3])
                                cells_consumed = 4
                        else:
                            cells_consumed = 1
                        
                        print(f"\n         COURT {court_in_row} (Court #{court_num}):")
                        print(f"            Court No (Coram): '{court_no}'")
                        print(f"            Running Item No: '{running_item}'")
                        print(f"            Case No: '{case_no}'")
                        print(f"            Passed Over Cases: '{passed_over}'")
                        print(f"            Cells consumed: {cells_consumed}")
                        
                        # MODIFIED: Extract ALL courts regardless of content
                        # No filtering for NS or COURT SESSION ENDED
                        court_data = {
                            "Court No (Coram)": court_no if court_no else f"Court {court_num}",
                            "Running Item No": running_item if running_item else "",
                            "Case No": case_no if case_no else "",
                            "Passed Over Cases": passed_over if passed_over else "",
                            "DateTime": scrape_time
                        }
                        all_courts_data.append(court_data)
                        print(f"            ✓ EXTRACTED (ALL courts extracted)")
                        
                        idx += cells_consumed  # Move to next court
                    else:
                        idx += 1  # Move to next cell
        
        print(f"\n{'='*90}")
        print(f"EXTRACTION SUMMARY:")
        print(f"{'='*90}")
        print(f"   ✓ Total courts extracted: {len(all_courts_data)}")
        print(f"   ✓ Timestamp: {scrape_time}")
        
        if all_courts_data:
            print(f"\n   Sample extracted data:")
            for i, court in enumerate(all_courts_data[:3], 1):
                print(f"      {i}. {court['Court No (Coram)']} | {court['Running Item No']} | {court['Case No']}")
        
        print(f"{'='*90}\n")
        
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
        df = df[["Court No (Coram)", "Running Item No", "Case No", "Passed Over Cases", "DateTime"]]
        
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
    print("=" * 90)
    print(" " * 25 + "TELANGANA HIGH COURT DISPLAY BOARD SCRAPER")
    print(" " * 30 + "(EXTRACT ALL COURTS VERSION)")
    print("=" * 90)
    print(f"URL: {URL}")
    print(f"Scrape Interval: {SCRAPE_INTERVAL} seconds")
    print(f"Save Location: {BASE_FOLDER}")
    print("=" * 90)
    
    # Create folder
    excel_path = create_folder()
    print(f"✓ Excel file path: {excel_path}")
    print("=" * 90)
    
    # Initialize driver
    print("\nInitializing Chrome driver...")
    driver = setup_driver()
    print("✓ Browser opened")
    print("=" * 90)
    
    cycle_count = 0
    first_cycle = True
    
    try:
        while True:
            cycle_count += 1
            current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            print(f"\n{'='*90}")
            print(f"CYCLE {cycle_count} - {current_time}")
            print(f"{'='*90}")
            
            # Scrape
            courts_data = scrape_display_board(driver)
            
            # Save to Excel
            if courts_data:
                success = save_to_excel(courts_data, excel_path, open_file=first_cycle)
                
                if success:
                    print(f"\n{'='*90}")
                    print(f"✓✓✓ CYCLE {cycle_count} COMPLETED SUCCESSFULLY ✓✓✓")
                    print(f"{'='*90}")
                    first_cycle = False
                else:
                    print(f"\n   ⚠ Save failed in cycle {cycle_count}")
            else:
                print(f"\n   ✗ No data scraped in cycle {cycle_count}")
            
            # Wait
            next_time = datetime.fromtimestamp(time.time() + SCRAPE_INTERVAL).strftime('%Y-%m-%d %H:%M:%S')
            print(f"\n{'─'*90}")
            print(f"⏳ Waiting {SCRAPE_INTERVAL} seconds")
            print(f"   Next cycle: {next_time}")
            print(f"{'─'*90}")
            time.sleep(SCRAPE_INTERVAL)
    
    except KeyboardInterrupt:
        print("\n" + "=" * 90)
        print("⚠ Script stopped by user")
        print(f"Total cycles: {cycle_count}")
        print("=" * 90)
    
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