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
        # Launch browser with more realistic settings - using slower timeout and slower launch
        print("Launching browser...")
        browser = None
        try:
            browser = p.chromium.launch(
                headless=False,
                slow_mo=100,  # Add a small delay between actions
                timeout=60000,  # Increase timeout to 60 seconds
                args=[
                    '--disable-blink-features=AutomationControlled',
                    '--disable-features=IsolateOrigins,site-per-process',
                    '--disable-site-isolation-trials',
                    '--disable-dev-shm-usage',  # Added for stability
                    '--no-sandbox'  # Added for stability
                ]
            )
            
            # Create a context with more realistic settings
            print("Creating browser context...")
            context = browser.new_context(
                viewport={'width': 1920, 'height': 1080},
                user_agent='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
                accept_downloads=True
            )
            
            # Set up download path with card name subfolder
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
            
            print("Creating new page...")
            page = context.new_page()

            try:
                # Navigate to American Express login page
                print("\nNavigating to American Express login page...")
                page.goto("https://www.americanexpress.com/en-us/account/login/", wait_until='networkidle', timeout=60000)
                time.sleep(3)

                # Fill login form
                print("Filling login form...")
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
                except Exception as e:
                    print(f"Verification method selection not needed or failed: {e}")
                
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
                    
                    # Select card based on parameter or default to Gold Card
                    card_to_select = card_name or "American Express Gold Card"
                    print(f"\nSelecting card: {card_to_select}...")
                    page.click("[role='combobox']")
                    time.sleep(3)
                    try:
                        page.click(f"text='{card_to_select}'")
                        print(f"Selected card: {card_to_select}")
                    except Exception as e:
                        print(f"Error selecting card '{card_to_select}': {e}")
                        print("Attempting to find card in dropdown...")
                        
                        # Try clicking dropdown and then check available cards
                        dropdown = page.query_selector("[role='combobox']")
                        if dropdown:
                            dropdown.click()
                            time.sleep(2)
                            # Get all available card options
                            card_options = page.query_selector_all("[role='option']")
                            print(f"Found {len(card_options)} card options")
                            
                            # Print available cards
                            available_cards = []
                            for option in card_options:
                                card_text = option.inner_text()
                                available_cards.append(card_text)
                            
                            if available_cards:
                                print(f"Available cards: {available_cards}")
                                # Try to find a partial match
                                for option in card_options:
                                    card_text = option.inner_text()
                                    if card_name.lower() in card_text.lower():
                                        print(f"Found partial match: {card_text}")
                                        option.click()
                                        print(f"Selected card: {card_text}")
                                        break
                            else:
                                print("No card options found in dropdown")
                                # Just select the first card
                                if card_options:
                                    first_card = card_options[0].inner_text()
                                    card_options[0].click()
                                    print(f"Selected first available card: {first_card}")
                        else:
                            print("Could not find card dropdown")
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
                    
                    # Take screenshot before clicking download button
                    screenshots_dir = os.path.join(os.path.dirname(__file__), 'screenshots')
                    os.makedirs(screenshots_dir, exist_ok=True)
                    page.screenshot(path=os.path.join(screenshots_dir, "before_first_download_click.png"))
                    
                    # Click first download button to open dialog
                    print("\nClicking download button to open dialog...")
                    try:
                        download_button = page.wait_for_selector("button:has-text('Download')")
                        if download_button:
                            print(f"Found download button. Attempting to click...")
                            download_button.click()
                            print("Clicked first download button successfully")
                        else:
                            print("Could not find download button with text 'Download'")
                    except Exception as e:
                        print(f"Error finding or clicking download button: {e}")
                        print("Trying alternative selectors for download button")
                        
                        # Try alternative selectors
                        alt_selectors = [
                            "[data-testid*='download']",
                            ".download-button",
                            "button.axp-activity__cta--download",
                            "button[aria-label*='download']",
                            # Try any button that might be the download button
                            "button.btn-primary",
                            "button.btn-secondary"
                        ]
                        
                        for selector in alt_selectors:
                            try:
                                btn = page.wait_for_selector(selector, timeout=3000)
                                if btn:
                                    print(f"Found button with selector: {selector}")
                                    btn.click()
                                    print(f"Clicked button with selector: {selector}")
                                    break
                            except:
                                print(f"Could not find or click button with selector: {selector}")
                    
                    time.sleep(3)
                    
                    # Take screenshot after clicking download button
                    page.screenshot(path=os.path.join(screenshots_dir, "after_first_download_click.png"))
                    
                    # Now handle the dialog that appears
                    print("\nHandling download dialog...")
                    
                    # Wait for dialog to appear and become stable
                    time.sleep(3)
                    
                    # Take screenshot for debugging
                    screenshots_dir = os.path.join(os.path.dirname(__file__), 'screenshots')
                    os.makedirs(screenshots_dir, exist_ok=True)
                    page.screenshot(path=os.path.join(screenshots_dir, "dialog_before_click.png"))
                    
                    # Define download handlers to capture downloads
                    download_started = False
                    
                    def handle_download(download):
                        nonlocal download_started
                        print(f"\n*** Download started: {download.suggested_filename} ***")
                        # Create more informative filename with date and card name
                        current_date = datetime.datetime.now().strftime("%Y%m%d")
                        card_identifier = card_name.replace(' ', '_').replace('/', '_').replace('\\', '_') if card_name else "AmexCard"
                        
                        filename_parts = download.suggested_filename.split('.')
                        if len(filename_parts) > 1:
                            ext = filename_parts[-1]
                            base_name = '.'.join(filename_parts[:-1])
                            new_filename = f"{base_name}_{card_identifier}_{current_date}.{ext}"
                        else:
                            new_filename = f"{download.suggested_filename}_{card_identifier}_{current_date}"
                        
                        download_path = os.path.join(download_dir, new_filename)
                        download.save_as(download_path)
                        print(f"Saved file to: {download_path}")
                        download_started = True
                    
                    # Set up download handler
                    page.on("download", handle_download)
                    
                    # Try different selectors for the download button in the dialog
                    print("\nLooking for download button in dialog...")
                    dialog_download_clicked = False
                    
                    # Take a screenshot of the dialog
                    dialog_screenshot = os.path.join(screenshots_dir, "download_dialog.png")
                    page.screenshot(path=dialog_screenshot)
                    print(f"Took screenshot of dialog: {dialog_screenshot}")
                    
                    # Expand selectors with more options
                    dialog_download_selectors = [
                        # API links
                        "a[href*='/api/servicing/v1/financials/documents']",
                        "a[href*='download']",
                        "a[href*='statements']",
                        "a[href*='activity']",
                        
                        # By attributes
                        "a[title='Download']",
                        "button[title='Download']",
                        "[data-test-id='axp-activity-download-footer-download-confirm']",
                        "[data-test-id*='download']",
                        "[data-testid*='download']",
                        
                        # By text content
                        "span:has-text('Download')",
                        "div:has-text('Download')",
                        "button:has-text('Download')",
                        "a:has-text('Download')",
                        
                        # By CSS classes
                        ".css-zmpgl6",
                        "a.btnStyle_lajeg_l",
                        "[class*='download']",
                        "[class*='btn']",
                        
                        # By container
                        ".modal button:has-text('Download')",
                        ".modal-footer button:has-text('Download')",
                        "[role='dialog'] button:has-text('Download')",
                        "[role='dialog'] a:has-text('Download')",
                        "[role='dialog'] span:has-text('Download')",
                        "[role='dialog'] .download",
                        
                        # Generic buttons that might be the download button
                        "[role='dialog'] button",
                        ".modal-footer button",
                        ".modal button"
                    ]
                    
                    # Try all selectors
                    for selector in dialog_download_selectors:
                        try:
                            print(f"Trying selector: {selector}")
                            elements = page.query_selector_all(selector)
                            print(f"Found {len(elements)} elements matching {selector}")
                            
                            # Try each element found
                            for i, element in enumerate(elements):
                                try:
                                    # Check if it's visible
                                    if element.is_visible():
                                        print(f"Element {i} is visible. Attempting to click...")
                                        element.click()
                                        print(f"Clicked element {i} with selector: {selector}")
                                        dialog_download_clicked = True
                                        time.sleep(5)  # Wait for download to start
                                        break
                                except Exception as e:
                                    print(f"Failed to click element {i}: {e}")
                            
                            if dialog_download_clicked:
                                break
                        except Exception as e:
                            print(f"Failed with selector {selector}: {e}")
                    
                    # If no selector worked, try specific approaches for the download dialog button
                    if not dialog_download_clicked:
                        print("\nTrying specific approaches for the download dialog button...")
                        
                        # Take another screenshot to help with debugging
                        page.screenshot(path=os.path.join(screenshots_dir, "before_special_approaches.png"))
                        
                        # Try method 1: Looking specifically for the blue button in the dialog
                        # Based on the screenshot, this is most likely to work
                        try:
                            print("Trying to click the blue Download button in the dialog...")
                            # This is likely to be the blue Download button in the modal
                            modal_buttons = page.query_selector_all("[role='dialog'] button")
                            print(f"Found {len(modal_buttons)} buttons in the dialog")
                            
                            # Try the right-most button which is typically the confirmation button
                            if len(modal_buttons) >= 2:
                                print("Clicking the right-most button (likely the Download button)")
                                modal_buttons[-1].click()
                                print("Clicked the right-most button in the dialog")
                                time.sleep(5)  # Wait to see if download starts
                            elif len(modal_buttons) == 1:
                                print("Only one button found, clicking it")
                                modal_buttons[0].click()
                                time.sleep(5)  # Wait to see if download starts
                        except Exception as e:
                            print(f"Failed to click modal button: {e}")
                        
                        # Try method 2: Direct coordinates for the blue Download button
                        # Based on the exact coordinates from screenshot
                        if not download_started:
                            try:
                                print("Trying exact coordinates for blue Download button...")
                                # These coordinates are for the blue Download button in the bottom right of the dialog
                                # Based on the screenshot you provided
                                page.mouse.click(800, 564)  # Adjusted based on the screenshot
                                print("Clicked at coordinates for blue Download button")
                                time.sleep(5)  # Wait to see if download starts
                            except Exception as e:
                                print(f"Failed to click at Download button coordinates: {e}")
                        
                        # Method 3: Try to click on the text "Download" in the blue button
                        if not download_started:
                            try:
                                print("Trying to click on 'Download' text...")
                                # Use evaluate to find and click on text content
                                page.evaluate("""
                                    (() => {
                                        const elements = Array.from(document.querySelectorAll('*'));
                                        for (const el of elements) {
                                            if (el.textContent.trim() === 'Download' && 
                                                window.getComputedStyle(el).backgroundColor.includes('rgb(0, 0')) {
                                                el.click();
                                                return true;
                                            }
                                        }
                                        return false;
                                    })();
                                """)
                                print("Attempted to click on 'Download' text through JavaScript")
                                time.sleep(5)  # Wait to see if download starts
                            except Exception as e:
                                print(f"Failed to click through JavaScript: {e}")
                        
                        # Method 4: Try a grid of coordinates around the button area
                        if not download_started:
                            print("Trying a grid around the Download button area...")
                            # Create a grid of coordinates centered on the Download button
                            center_x, center_y = 800, 564  # From the screenshot
                            for x_offset in [-20, 0, 20]:
                                for y_offset in [-10, 0, 10]:
                                    try:
                                        x, y = center_x + x_offset, center_y + y_offset
                                        print(f"Clicking at ({x}, {y})...")
                                        page.mouse.click(x, y)
                                        time.sleep(3)
                                        if download_started:
                                            print(f"Success! Coordinates ({x}, {y}) worked.")
                                            break
                                    except Exception as e:
                                        print(f"Failed to click at ({x}, {y}): {e}")
                                    
                                if download_started:
                                    break
                        
                        # Take screenshot after trying all approaches
                        page.screenshot(path=os.path.join(screenshots_dir, "after_special_approaches.png"))
                    
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
                        print(f"Summary:")
                        print(f"- Card: {card_name or 'American Express Gold Card'}")
                        print(f"- Download location: {download_dir}")
                        
                        # Write a summary to a log file
                        log_dir = os.path.join(os.path.dirname(__file__), 'logs')
                        os.makedirs(log_dir, exist_ok=True)
                        log_file = os.path.join(log_dir, f"download_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
                        
                        try:
                            with open(log_file, 'w') as f:
                                f.write(f"Download Summary\n")
                                f.write(f"---------------\n")
                                f.write(f"Date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                                f.write(f"Card: {card_name or 'American Express Gold Card'}\n")
                                f.write(f"Download location: {download_dir}\n")
                                f.write(f"Download successful: Yes\n")
                            print(f"Download log saved to: {log_file}")
                        except Exception as e:
                            print(f"Could not write log file: {e}")
                    else:
                        print("\nDownload did not start within the timeout period.")
                        
                        # Write a failure log
                        log_dir = os.path.join(os.path.dirname(__file__), 'logs')
                        os.makedirs(log_dir, exist_ok=True)
                        log_file = os.path.join(log_dir, f"download_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_FAILED.txt")
                        
                        try:
                            with open(log_file, 'w') as f:
                                f.write(f"Download Summary\n")
                                f.write(f"---------------\n")
                                f.write(f"Date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                                f.write(f"Card: {card_name or 'American Express Gold Card'}\n")
                                f.write(f"Download location: {download_dir}\n")
                                f.write(f"Download successful: No\n")
                                f.write(f"Reason: Timeout waiting for download to start\n")
                            print(f"Failure log saved to: {log_file}")
                        except Exception as e:
                            print(f"Could not write log file: {e}")
                        
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
                screenshots_dir = os.path.join(os.path.dirname(__file__), 'screenshots')
                os.makedirs(screenshots_dir, exist_ok=True)
                try:
                    page.screenshot(path=os.path.join(screenshots_dir, "error_screenshot.png"))
                    print(f"Error screenshot saved to {os.path.join(screenshots_dir, 'error_screenshot.png')}")
                except:
                    print("Could not take error screenshot")
            
            finally:
                # Close browser
                time.sleep(3)  # Wait before closing
                try:
                    if browser:
                        browser.close()
                except Exception as e:
                    print(f"Error closing browser: {e}")
                    
        except Exception as e:
            print(f"Error launching browser: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Download American Express statements')
    parser.add_argument('--card', type=str, help='Card name to select (e.g., "American Express Gold Card", "Platinum Card")')
    args = parser.parse_args()
    
    main(card_name=args.card)