#!/usr/bin/env python
"""
Script to download American Express statements from the Statements and Year End Summaries section.
"""
from playwright.sync_api import sync_playwright
import time
import os
from dotenv import load_dotenv
import datetime
import sys
import argparse

def main(card_name=None):
    # Load environment variables
    env_path = os.path.join(os.path.dirname(__file__), 'config', '.env')
    load_dotenv(env_path)
    
    # Get credentials and verify they're loaded
    amex_username = os.getenv('AMEX_USERNAME')
    amex_password = os.getenv('AMEX_PASSWORD')
    
    if not amex_username or not amex_password:
        print("\nError: Could not load credentials from .env file")
        return
    
    print(f"\nCredentials loaded successfully")
    
    # Create directories for logs and screenshots
    logs_dir = os.path.join(os.path.dirname(__file__), 'logs')
    screenshots_dir = os.path.join(os.path.dirname(__file__), 'screenshots')
    os.makedirs(logs_dir, exist_ok=True)
    os.makedirs(screenshots_dir, exist_ok=True)
    
    # Set up log file
    current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(logs_dir, f"download_log_{current_time}.txt")
    
    # Clear OTP file before starting
    icloud_path = "/Users/sahil/Library/Mobile Documents/com~apple~CloudDocs/OTP/otp.txt"
    try:
        with open(icloud_path, 'w') as f:
            pass
        print("Cleared OTP file before starting")
    except Exception as e:
        print(f"Could not clear OTP file: {e}")

    # Initialize browser
    with sync_playwright() as p:
        # Launch browser with more realistic settings
        browser = p.chromium.launch(
            headless=False,
            args=[
                '--disable-blink-features=AutomationControlled',
                '--disable-features=IsolateOrigins,site-per-process',
                '--disable-site-isolation-trials'
            ]
        )
        
        # Create a context with more realistic settings
        context = browser.new_context(
            viewport={'width': 1920, 'height': 1080},
            user_agent='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
            accept_downloads=True
        )
        
        # Set up download path
        base_download_dir = os.path.expanduser("~/Downloads/AmexStatements")
        os.makedirs(base_download_dir, exist_ok=True)
        
        # Create card-specific subfolder if card name is provided
        if card_name:
            # Convert card name to filesystem-friendly format
            card_folder_name = card_name.replace(' ', '_').replace('/', '_').replace('\\', '_')
            download_dir = os.path.join(base_download_dir, card_folder_name)
            os.makedirs(download_dir, exist_ok=True)
            print(f"Downloads will be saved to {download_dir}")
        else:
            download_dir = base_download_dir
            print(f"Downloads will be saved to {download_dir} (default location)")
        
        page = context.new_page()

        try:
            # Navigate to American Express login page
            print("\nNavigating to American Express login page...")
            page.goto("https://www.americanexpress.com/en-us/account/login/", wait_until='networkidle')
            time.sleep(3)
            
            # Take a screenshot of the login page
            page.screenshot(path=os.path.join(screenshots_dir, "login_page.png"))

            # Fill login form
            print("Filling login credentials...")
            page.fill("#eliloUserID", amex_username)
            page.fill("#eliloPassword", amex_password)
            page.click("#loginSubmit")
            time.sleep(5)
            
            # Handle two-step verification
            print("\nHandling two-step verification...")
            try:
                # Try to find and click change verification method
                change_button = page.wait_for_selector("button:has-text('Change verification method')", timeout=5000)
                if change_button:
                    change_button.click()
                    time.sleep(2)
                    
                    # Try SMS first, if not available use email
                    try:
                        sms_option = page.wait_for_selector("button:has-text('One-time password (SMS)')", timeout=3000)
                        if sms_option and sms_option.is_visible():
                            sms_option.click()
                            print("Selected SMS verification")
                    except:
                        # If SMS not available, try email
                        email_option = page.wait_for_selector("button:has-text('Email')", timeout=3000)
                        if email_option:
                            email_option.click()
                            print("Selected Email verification (SMS not available)")
                    
                    time.sleep(2)
            except:
                print("Verification method selection not needed")
            
            # Get OTP code
            print("\nWaiting for OTP file...")
            otp_code = None
            max_attempts = 6
            attempt = 0
            
            while attempt < max_attempts and not otp_code:
                attempt += 1
                print(f"Waiting for OTP... attempt {attempt}/{max_attempts}")
                time.sleep(10)
                
                try:
                    with open(icloud_path, 'r') as f:
                        lines = [line.strip() for line in f.readlines() if line.strip()]
                        if lines:
                            otp_code = lines[-1]
                            print(f"OTP found on attempt {attempt}")
                except Exception as e:
                    if attempt == max_attempts:
                        print(f"Could not read OTP after {max_attempts} attempts: {e}")
                        return
            
            if not otp_code:
                print("OTP code not found or empty.")
                return
                
            print(f"OTP code read: {otp_code}")
            
            # Enter OTP
            page.fill("input[type='text']", otp_code)
            page.click("button:has-text('Verify')")
            time.sleep(5)
            page.click("button:has-text('Continue')")
            time.sleep(5)
            
            # Clear OTP file
            with open(icloud_path, 'w') as f:
                pass
            
            # Take a screenshot after login
            page.screenshot(path=os.path.join(screenshots_dir, "after_login.png"))
            
            # Select card based on parameter or default to Gold Card
            card_to_select = card_name or "American Express Gold Card"
            print(f"\nSelecting card: {card_to_select}...")
            
            # Click on card selector dropdown
            try:
                page.click("[role='combobox']")
                time.sleep(3)
                print("Clicked on card dropdown")
                
                # Try to find and select the specified card
                try:
                    card_option = page.wait_for_selector(f"text='{card_to_select}'", timeout=5000)
                    if card_option:
                        card_option.click()
                        print(f"Selected card: {card_to_select}")
                        time.sleep(5)
                except Exception as e:
                    print(f"Could not automatically select card: {e}")
                    print("Please manually select the desired card")
                    time.sleep(10)  # Give user time to select card manually
            except Exception as e:
                print(f"Error clicking card dropdown: {e}")
                print("If you're already on the correct card, that's fine.")
            
            # Take a screenshot after card selection
            page.screenshot(path=os.path.join(screenshots_dir, "after_card_selection.png"))
            
            # Navigate to Statements & Activity
            print("\nNavigating to Statements & Activity...")
            try:
                statements_link = page.wait_for_selector("span:has-text('Statements & Activity')", timeout=5000)
                if statements_link:
                    statements_link.click()
                    print("Clicked on Statements & Activity")
                    time.sleep(5)
            except Exception as e:
                print(f"Error clicking Statements & Activity: {e}")
                print("Attempting alternative navigation")
                
                # Try alternative navigation
                try:
                    page.goto("https://www.americanexpress.com/en-us/account/statements", wait_until='networkidle', timeout=30000)
                    print("Navigated directly to statements page")
                    time.sleep(5)
                except Exception as e2:
                    print(f"Direct navigation failed: {e2}")
                    print("Please navigate to Statements & Activity manually")
                    time.sleep(15)  # Give user time to navigate manually
            
            # Take a screenshot of the Statements & Activity page
            page.screenshot(path=os.path.join(screenshots_dir, "statements_activity_page.png"))
            
            # NEXT STEP: Click on "Statements and Year End Summaries"
            print("\nNavigating to Statements and Year End Summaries...")
            try:
                # Try to find and click on the "Statements and Year End Summaries" link
                summaries_link = page.wait_for_selector("text='Statements and Year End Summaries'", timeout=5000)
                if summaries_link:
                    summaries_link.click()
                    print("Clicked on Statements and Year End Summaries")
                    time.sleep(5)
            except Exception as e:
                print(f"Error clicking Statements and Year End Summaries: {e}")
                print("Please click on Statements and Year End Summaries manually")
                time.sleep(15)  # Give user time to click manually
            
            # Take a screenshot after navigating to Statements and Year End Summaries
            page.screenshot(path=os.path.join(screenshots_dir, "statements_summaries_page.png"))
            
            print("\nSuccessfully navigated to Statements and Year End Summaries section!")
            
            # Click on "Older Statements" to view more statements
            print("\nExpanding to view Older Statements...")
            try:
                # Try to find and click on the "Older Statements" dropdown
                older_statements_button = page.wait_for_selector("text='Older Statements'", timeout=5000)
                if older_statements_button:
                    older_statements_button.click()
                    print("Clicked on Older Statements")
                    time.sleep(5)
                    
                    # Take a screenshot after expanding Older Statements
                    page.screenshot(path=os.path.join(screenshots_dir, "older_statements_expanded.png"))
                    print("Older Statements section expanded successfully")
            except Exception as e:
                print(f"Error expanding Older Statements: {e}")
                print("Could not expand Older Statements section automatically")
            
            # Now click on each download link directly through the UI
            print("\nPreparing to download statements...")
            try:
                # Create a directory for downloaded statements
                statements_dir = os.path.join(download_dir, "statements")
                os.makedirs(statements_dir, exist_ok=True)
                print(f"Created directory: {statements_dir}")
                
                # Set up download event handler
                download_count = 0
                current_date = ""
                
                def handle_download(download):
                    nonlocal download_count, current_date
                    # Get suggested filename
                    suggested_filename = download.suggested_filename
                    
                    # Create a filename with the statement date if available
                    if current_date:
                        date_str = current_date.replace(' ', '_').replace(',', '')
                        # Use the appropriate extension from the suggested filename or default to csv
                        ext = os.path.splitext(suggested_filename)[1] if suggested_filename and '.' in suggested_filename else '.csv'
                        filename = f"AmexStatement_{date_str}{ext}"
                    else:
                        # Fallback to generic name with timestamp
                        ext = os.path.splitext(suggested_filename)[1] if suggested_filename and '.' in suggested_filename else '.csv'
                        filename = f"AmexStatement_{download_count+1}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
                    
                    # Save the file
                    filepath = os.path.join(statements_dir, filename)
                    download.save_as(filepath)
                    print(f"Downloaded statement to: {filepath}")
                    download_count += 1
                
                # Register the download handler
                page.on("download", handle_download)
                
                # Find all download buttons
                download_buttons = page.query_selector_all("button:has-text('Download'), a:has-text('Download')")
                print(f"Found {len(download_buttons)} download buttons")
                
                # Get all date elements for pairing with download buttons
                date_elements = page.query_selector_all("td:first-child, .date, [data-closing-date], .statement-date")
                print(f"Found {len(date_elements)} date elements")
                
                # Process the Recent Statements section first
                print("\nDownloading Recent Statements...")
                # Get all the rows in the recent statements table
                recent_rows = page.query_selector_all("table tr, .statement-row")
                
                for row in recent_rows:
                    try:
                        # Skip header rows
                        if row.get_attribute("role") == "heading" or "header" in row.get_attribute("class", ""):
                            continue
                        
                        # Try to get the date from the row
                        date_element = row.query_selector("td:first-child")
                        if date_element:
                            current_date = date_element.inner_text().strip()
                            print(f"Found date: {current_date}")
                        
                        # Find the download button in this row
                        dl_button = row.query_selector("button:has-text('Download'), a:has-text('Download')")
                        if dl_button:
                            print(f"Downloading statement for {current_date}...")
                            dl_button.click()
                            print(f"Clicked download button for {current_date}")
                            
                            # Wait for the file type selection dialog to appear
                            time.sleep(2)
                            
                            # Take a screenshot of the file type selection dialog
                            dialog_screenshot_path = os.path.join(screenshots_dir, f"file_type_dialog_{download_count+1}.png")
                            page.screenshot(path=dialog_screenshot_path)
                            print(f"Took screenshot of file type dialog: {dialog_screenshot_path}")
                            
                            # Select CSV option
                            try:
                                # Look for CSV radio button
                                csv_option = page.wait_for_selector("input[type='radio'][id*='csv'], label:has-text('CSV')", timeout=5000)
                                if csv_option:
                                    csv_option.click()
                                    print("Selected CSV format")
                                    time.sleep(1)
                                    
                                    # Take a screenshot after selecting CSV
                                    csv_selected_path = os.path.join(screenshots_dir, f"csv_selected_{download_count+1}.png")
                                    page.screenshot(path=csv_selected_path)
                                    print(f"Took screenshot after selecting CSV: {csv_selected_path}")
                                    
                                    # Click the Download button in the dialog
                                    download_button = page.wait_for_selector("button:has-text('Download'):not([disabled])", timeout=5000)
                                    if download_button:
                                        download_button.click()
                                        print("Clicked Download button in the dialog")
                                    else:
                                        print("Could not find Download button in the dialog")
                                else:
                                    print("Could not find CSV option")
                            except Exception as e:
                                print(f"Error selecting CSV format: {e}")
                            
                            # Wait for download to complete
                            time.sleep(5)
                    except Exception as e:
                        print(f"Error processing row: {e}")
                
                # If there are no clear rows, just click on all download buttons sequentially
                if download_count == 0:
                    print("\nFalling back to sequential download of all buttons...")
                    for i, button in enumerate(download_buttons):
                        try:
                            # Try to find a nearby date element
                            current_date = f"Statement_{i+1}"
                            print(f"Downloading statement {i+1}...")
                            button.click()
                            print(f"Clicked download button {i+1}")
                            
                            # Wait for the file type selection dialog to appear
                            time.sleep(2)
                            
                            # Take a screenshot of the file type selection dialog
                            dialog_screenshot_path = os.path.join(screenshots_dir, f"file_type_dialog_{i+1}.png")
                            page.screenshot(path=dialog_screenshot_path)
                            print(f"Took screenshot of file type dialog: {dialog_screenshot_path}")
                            
                            # Select CSV option
                            try:
                                # Look for CSV radio button
                                csv_option = page.wait_for_selector("input[type='radio'][id*='csv'], label:has-text('CSV')", timeout=5000)
                                if csv_option:
                                    csv_option.click()
                                    print("Selected CSV format")
                                    time.sleep(1)
                                    
                                    # Take a screenshot after selecting CSV
                                    csv_selected_path = os.path.join(screenshots_dir, f"csv_selected_{i+1}.png")
                                    page.screenshot(path=csv_selected_path)
                                    print(f"Took screenshot after selecting CSV: {csv_selected_path}")
                                    
                                    # Click the Download button in the dialog
                                    download_button = page.wait_for_selector("button:has-text('Download'):not([disabled])", timeout=5000)
                                    if download_button:
                                        download_button.click()
                                        print("Clicked Download button in the dialog")
                                    else:
                                        print("Could not find Download button in the dialog")
                                else:
                                    print("Could not find CSV option")
                            except Exception as e:
                                print(f"Error selecting CSV format: {e}")
                            
                            # Wait for download to complete
                            time.sleep(5)
                        except Exception as e:
                            print(f"Error clicking download button {i+1}: {e}")
                
                print(f"\nCompleted downloading {download_count} statements to {statements_dir}")
            
            except Exception as e:
                print(f"Error downloading statements: {e}")
                print("Could not automatically download statements")
            
            # Take a final screenshot of the Statements and Year End Summaries page
            final_screenshot_path = os.path.join(screenshots_dir, "final_statements_page.png")
            page.screenshot(path=final_screenshot_path)
            print(f"Final screenshot saved to: {final_screenshot_path}")
            
            # Wait 30 seconds before asking to close browser
            print("\nWaiting 30 seconds to examine the page...")
            time.sleep(30)
            
            # Wait for user to decide whether to keep browser open
            try:
                keep_open = input("\nDo you want to keep the browser open? (y/n): ")
                if keep_open.lower() != 'y':
                    browser.close()
                    print("Browser closed")
                else:
                    print("Browser left open - please close it manually when finished")
            except:
                # If running in an environment where input is not possible
                print("\nNo input detected - waiting additional 30 seconds before closing browser")
                time.sleep(30)
                # Take one more screenshot before closing
                page.screenshot(path=os.path.join(screenshots_dir, "final_statements_page_before_close.png"))
                print("Final screenshot taken before closing browser")
            
            # Log successful completion
            with open(log_file, 'w') as f:
                f.write("Download Summary\n")
                f.write("---------------\n")
                f.write(f"Date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Card: {card_name or 'Default Card'}\n")
                f.write(f"Navigation: Completed up to Statements and Year End Summaries\n")
                f.write(f"Status: Partial completion - next steps to be implemented\n")
        
        except Exception as e:
            print(f"\nError during process: {e}")
            # Take screenshot on error
            page.screenshot(path=os.path.join(screenshots_dir, f"error_screenshot_{current_time}.png"))
            
            # Log error
            with open(log_file, 'w') as f:
                f.write("Download Summary\n")
                f.write("---------------\n")
                f.write(f"Date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Card: {card_name or 'Default Card'}\n")
                f.write(f"Error: {str(e)}\n")
                f.write(f"Status: Failed\n")
        
        finally:
            # Close browser if still open
            if 'keep_open' not in locals() or keep_open.lower() != 'y':
                try:
                    browser.close()
                    print("Browser closed")
                except:
                    pass

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Download American Express statements from Year End Summaries')
    parser.add_argument('--card', type=str, help='Card name to select (e.g., "American Express Gold Card", "Platinum Card")')
    args = parser.parse_args()
    
    main(card_name=args.card)