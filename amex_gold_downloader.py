from playwright.sync_api import sync_playwright
import time
import os
from dotenv import load_dotenv
import datetime
import sys

def main():
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
        download_dir = os.path.expanduser("~/Downloads/AmexStatements")
        os.makedirs(download_dir, exist_ok=True)
        
        page = context.new_page()

        try:
            # Navigate to American Express login page
            print("\nNavigating to American Express login page...")
            page.goto("https://www.americanexpress.com/en-us/account/login/", wait_until='networkidle')
            time.sleep(3)

            # Fill login form
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
            
            if otp_code:
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
                
                # Select Personal Gold Card
                print("\nSelecting Personal Gold Card...")
                page.click("[role='combobox']")
                time.sleep(3)
                page.click("text='American Express Gold Card'")
                time.sleep(5)
                
                # Navigate to Statements & Activity
                print("\nNavigating to Statements & Activity...")
                page.click("span:has-text('Statements & Activity')")
                time.sleep(5)
                
                # Go to Custom Date Range
                print("\nNavigating to Custom Date Range...")
                page.click("a[href='/activity/search']")
                time.sleep(5)
                
                # Click search button (3rd one)
                print("\nClicking search button...")
                search_buttons = page.query_selector_all("button:has-text('Search'), [role='button']:has-text('Search')")
                if len(search_buttons) >= 3:
                    print(f"Found {len(search_buttons)} search buttons. Clicking the 3rd one...")
                    search_buttons[2].click()
                else:
                    print(f"Not enough search buttons found (found {len(search_buttons)}, need at least 3)")
                    # Fallback to clicking the last one
                    if search_buttons:
                        search_buttons[-1].click()
                
                # Wait for search results
                time.sleep(10)
                
                # Click first download button to open dialog
                print("\nClicking download button to open dialog...")
                download_button = page.wait_for_selector("button:has-text('Download')")
                download_button.click()
                time.sleep(3)
                
                # Now handle the dialog that appears
                print("\nHandling download dialog...")
                
                # Wait for dialog to appear and become stable
                time.sleep(3)
                
                # Take screenshot for debugging
                page.screenshot(path="dialog_before_click.png")
                
                # Define download handlers to capture downloads
                download_started = False
                
                def handle_download(download):
                    nonlocal download_started
                    print(f"\n*** Download started: {download.suggested_filename} ***")
                    download_path = os.path.join(download_dir, download.suggested_filename)
                    download.save_as(download_path)
                    print(f"Saved file to: {download_path}")
                    download_started = True
                
                # Set up download handler
                page.on("download", handle_download)
                
                # Try different selectors for the download button in the dialog
                print("\nLooking for download button in dialog...")
                dialog_download_clicked = False
                
                # Selectors based on the screenshot and HTML inspection
                dialog_download_selectors = [
                    "a[href*='/api/servicing/v1/financials/documents']",
                    "a[title='Download']",
                    "[data-test-id='axp-activity-download-footer-download-confirm']",
                    "span:has-text('Download')",
                    ".css-zmpgl6",
                    "a.btnStyle_lajeg_l",
                    ".modal button:has-text('Download')",
                    ".modal-footer button:has-text('Download')",
                    "[role='dialog'] button:has-text('Download')",
                    "[role='dialog'] a:has-text('Download')"
                ]
                
                for selector in dialog_download_selectors:
                    try:
                        print(f"Trying selector: {selector}")
                        button = page.wait_for_selector(selector, timeout=3000)
                        if button:
                            print(f"Found dialog download button with selector: {selector}")
                            button.click()
                            print("Clicked dialog download button")
                            dialog_download_clicked = True
                            time.sleep(5)  # Wait for download to start
                            break
                    except Exception as e:
                        print(f"Failed with selector {selector}: {e}")
                
                # If no selector worked, try coordinates based on the screenshot
                if not dialog_download_clicked:
                    print("\nTrying fixed coordinates for download button...")
                    page.mouse.click(650, 492)  # Adjust coordinates based on your dialog
                    time.sleep(5)
                
                # Wait for download to complete
                print("\nWaiting for download to complete...")
                wait_time = 0
                max_wait = 30  # seconds
                
                while not download_started and wait_time < max_wait:
                    time.sleep(1)
                    wait_time += 1
                    print(f"Waiting... {wait_time}/{max_wait} seconds")
                
                if download_started:
                    print("\nDownload completed successfully!")
                else:
                    print("\nDownload did not start within the timeout period.")
                    
                    # Check if any files were downloaded to the download directory
                    print("\nChecking download directory...")
                    files = os.listdir(download_dir)
                    if files:
                        print(f"Files in download directory: {files}")
                    else:
                        print("No files found in download directory.")
                        
                        # Check default downloads folder for recent Excel files
                        print("\nChecking default Downloads folder for recent files...")
                        downloads_folder = os.path.expanduser("~/Downloads")
                        recent_time = time.time() - 300  # Files in the last 5 minutes
                        
                        recent_files = []
                        for file in os.listdir(downloads_folder):
                            file_path = os.path.join(downloads_folder, file)
                            if os.path.getmtime(file_path) > recent_time and (file.endswith('.xlsx') or file.endswith('.xls') or 'amex' in file.lower()):
                                recent_files.append(file)
                        
                        if recent_files:
                            print(f"Recent files that might be the download: {recent_files}")
                        else:
                            print("No recent relevant files found in Downloads folder.")
                
                print("\nScript completed!")
            else:
                print("OTP code not found or empty.")
        
        except Exception as e:
            print(f"\nError during process: {e}")
            # Take screenshot on error
            page.screenshot(path="error_screenshot.png")
        
        finally:
            # Close browser
            time.sleep(3)  # Wait before closing
            browser.close()

if __name__ == "__main__":
    main() 