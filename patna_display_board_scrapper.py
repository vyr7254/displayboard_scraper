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
SCRAPE_INTERVAL = 30  # seconds (adjust as needed)
BASE_FOLDER = r"D:\CourtDisplayBoardScraper\displayboardexcel\patna_hc_excels"
EXCEL_FILE = "PatnaHC_DisplayBoard_Data.xlsx"

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


def extract_cell_content(cell):
    """
    Extract content from cell - either input button value or text content
    Uses innerHTML parsing for speed (like SC scraper)
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
            for i, court in enumerate(all_courts_data[:5], 1):
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
        df = df[["Court Number", "Item No", "Case Number", "Full Case Details", "DateTime"]]
        
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
    print(" " * 30 + "PATNA HIGH COURT DISPLAY BOARD SCRAPER")
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