
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
URL = "https://highcourt.cg.gov.in/hcbspcourtview/court1.php"
SCRAPE_INTERVAL = 30  # seconds
BASE_FOLDER = r"D:\CourtDisplayBoardScraper\displayboardexcel\chattisgarh_hc_excel"
EXCEL_FILE = "Chattisgarh_HighCourt_DisplayBoard_Data.xlsx"

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


# ==================== SCRAPING FUNCTIONS ====================

def scrape_display_board(driver):
    """
    Scrape courts from Chhattisgarh High Court display board
    Each row = 1 court with 6 columns: Court No. | List Type | Sl. No. | Case No. | Party | Purpose
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
        
        print("\n" + "="*100)
        print("ANALYZING PAGE STRUCTURE - EXTRACTING ALL COURTS...")
        print("="*100)
        
        # Find all tables
        tables = driver.find_elements(By.TAG_NAME, "table")
        print(f"   → Found {len(tables)} table(s) on the page")
        
        # Use the first table (or find the correct one based on attributes)
        table = tables[0]
        rows = table.find_elements(By.TAG_NAME, "tr")
        print(f"   → Found {len(rows)} total rows")
        
        # Check headers
        if len(rows) > 0:
            print(f"\n{'─'*100}")
            print("HEADER ROW:")
            print(f"{'─'*100}")
            header_cells = rows[0].find_elements(By.TAG_NAME, "th")
            if not header_cells:
                header_cells = rows[0].find_elements(By.TAG_NAME, "td")
            
            for idx, cell in enumerate(header_cells):
                header_text = extract_cell_text(cell)
                print(f"   Header[{idx}]: '{header_text}'")
        
        # Extract court data
        all_courts_data = []
        
        print(f"\n{'─'*100}")
        print("EXTRACTING DATA FROM ROWS:")
        print(f"{'─'*100}")
        
        # Process each row (skip header row)
        for row_idx, row in enumerate(rows[1:], 1):
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                
                # Skip empty rows or rows with insufficient cells
                if len(cells) < 6:
                    # Check if it's a "Court Not in Session" or similar message
                    if len(cells) > 0:
                        row_text = extract_cell_text(cells[0])
                        if "Court Not in Session" in row_text or "not in session" in row_text.lower():
                            court_no = extract_cell_text(cells[0]) if len(cells) > 0 else ""
                            print(f"\n   ROW {row_idx}: {row_text}")
                            
                            court_data = {
                                "Court No.": court_no,
                                "List Type": "Not in Session",
                                "Sl. No.": "",
                                "Case No.": "",
                                "Party": "",
                                "Purpose": "Court Not in Session",
                                "DateTime": scrape_time
                            }
                            all_courts_data.append(court_data)
                            print(f"      ✓ EXTRACTED (Not in Session)")
                    continue
                
                # Each row should have 6 columns
                court_no = extract_cell_text(cells[0])
                list_type = extract_cell_text(cells[1])
                sl_no = extract_cell_text(cells[2])
                case_no = extract_cell_text(cells[3])
                party = extract_cell_text(cells[4])
                purpose = extract_cell_text(cells[5])
                
                print(f"\n   ROW {row_idx} (Court {court_no}):")
                print(f"      Court No.: '{court_no}'")
                print(f"      List Type: '{list_type}'")
                print(f"      Sl. No.: '{sl_no}'")
                print(f"      Case No.: '{case_no}'")
                print(f"      Party: '{party}'")
                print(f"      Purpose: '{purpose}'")
                
                # Create court data dictionary
                court_data = {
                    "Court No.": court_no if court_no else "",
                    "List Type": list_type if list_type else "",
                    "Sl. No.": sl_no if sl_no else "",
                    "Case No.": case_no if case_no else "",
                    "Party": party if party else "",
                    "Purpose": purpose if purpose else "",
                    "DateTime": scrape_time
                }
                
                all_courts_data.append(court_data)
                print(f"      ✓ EXTRACTED")
                    
            except Exception as e:
                print(f"\n   ✗ Error processing row {row_idx}: {str(e)}")
                continue
        
        print(f"\n{'='*100}")
        print(f"EXTRACTION SUMMARY:")
        print(f"{'='*100}")
        print(f"   ✓ Total courts extracted: {len(all_courts_data)}")
        print(f"   ✓ Timestamp: {scrape_time}")
        
        if all_courts_data:
            print(f"\n   Sample extracted data:")
            for i, court in enumerate(all_courts_data[:3], 1):
                print(f"      {i}. Court {court['Court No.']} | {court['List Type']} | Sl. {court['Sl. No.']} | {court['Case No.']}")
        
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
        df = df[["Court No.", "List Type", "Sl. No.", "Case No.", "Party", "Purpose", "DateTime"]]
        
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
    print(" " * 30 + "CHHATTISGARH HIGH COURT DISPLAY BOARD SCRAPER")
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