#!/usr/bin/env python
"""
Specialized script for downloading Platinum Card statements.
This script addresses the specific download dialog issue with the Platinum Card.
"""
from playwright.sync_api import sync_playwright
import time
import os
import sys
from dotenv import load_dotenv
import datetime
import argparse

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

    # Create directories for logs and screenshots
    logs_dir = os.path.join(os.path.dirname(__file__), 'logs')
    screenshots_dir = os.path.join(os.path.dirname(__file__), 'screenshots')
    os.makedirs(logs_dir, exist_ok=True)
    os.makedirs(screenshots_dir, exist_ok=True)
    
    # Initialize browser
    with sync_playwright() as p:
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
            
            # Create card-specific subfolder
            card_folder_name = "Platinum_Card"
            download_dir = os.path.join(base_download_dir, card_folder_name)
            os.makedirs(download_dir, exist_ok=True)
            print(f"Downloads will be saved to {download_dir}")
            
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
                    
                    # Select Platinum Card
                    print("\nSelecting Platinum Card...")
                    
                    # Check if we're already on the dashboard
                    try:
                        current_card = page.query_selector(".card-name")
                        if current_card:
                            card_text = current_card.inner_text()
                            print(f"Current card displayed: {card_text}")
                            
                            if "platinum" in card_text.lower():
                                print("Already on Platinum Card - no need to select it")
                                card_selected = True
                            else:
                                print("Different card selected, need to change to Platinum Card")
                                # Click the dropdown to change cards
                                page.click("[role='combobox']")
                                time.sleep(3)
                        else:
                            # No card displayed yet, click the dropdown
                            page.click("[role='combobox']")
                            time.sleep(3)
                    except Exception as e:
                        print(f"Error checking current card: {e}")
                        # Try to click the dropdown anyway
                        try:
                            page.click("[role='combobox']")
                            time.sleep(3)
                        except:
                            print("Could not click card dropdown")
                    
                    # Try different selectors for Platinum Card
                    platinum_selectors = [
                        "text='Platinum Card'", 
                        "text='The Platinum Card'",
                        "text='Platinum Card®'",
                        "text='American Express Platinum Card'",
                        "text=Platinum",
                        "[alt='Platinum Card']",
                        "[alt*='Platinum']"
                    ]
                    
                    card_selected = False
                    for selector in platinum_selectors:
                        try:
                            print(f"Trying to select card with: {selector}")
                            card = page.wait_for_selector(selector, timeout=3000)
                            if card:
                                card.click()
                                print(f"Selected card with: {selector}")
                                card_selected = True
                                break
                        except Exception as e:
                            print(f"Failed to select with {selector}: {e}")
                    
                    if not card_selected:
                        print("Could not find Platinum Card by text. Trying to list and select available cards...")
                        # Take a screenshot of available cards
                        page.screenshot(path=os.path.join(screenshots_dir, "available_cards.png"))
                        
                        # Try to get all available options
                        try:
                            options = page.query_selector_all("[role='option']")
                            print(f"Found {len(options)} card options")
                            
                            for i, option in enumerate(options):
                                try:
                                    text = option.inner_text()
                                    print(f"Card option {i}: {text}")
                                    if "platinum" in text.lower():
                                        print(f"Found Platinum Card at index {i}")
                                        option.click()
                                        card_selected = True
                                        break
                                except Exception as e:
                                    print(f"Error reading option {i}: {e}")
                            
                            # If no Platinum Card, just click the first card
                            if not card_selected and len(options) > 0:
                                print("Platinum Card not found. Clicking first available card.")
                                options[0].click()
                                card_selected = True
                        except Exception as e:
                            print(f"Error getting card options: {e}")
                    
                    time.sleep(5)
                    
                    # Navigate to Statements & Activity
                    print("\nNavigating to Statements & Activity...")
                    try:
                        # Look for Statements & Activity in the main navigation
                        statements_link = page.wait_for_selector("span:has-text('Statements & Activity'), a:has-text('Statements & Activity')", timeout=5000)
                        if statements_link:
                            statements_link.click()
                            print("Clicked Statements & Activity link")
                            time.sleep(5)
                        else:
                            print("Could not find Statements & Activity link")
                    except Exception as e:
                        print(f"Error finding Statements & Activity link: {e}")
                        # Check if we're already on the statements page
                        if "statement" in page.url.lower() or "activity" in page.url.lower():
                            print("Already on Statements & Activity page")
                        else:
                            print("Attempting to navigate directly to activity search")
                            page.goto("https://www.americanexpress.com/en-us/account/activity/search", wait_until='networkidle', timeout=30000)
                    
                    time.sleep(5)
                    
                    # Take screenshot of the current page
                    page.screenshot(path=os.path.join(screenshots_dir, "before_search_page.png"))
                    
                    # Go to Custom Date Range
                    print("\nNavigating to Custom Date Range...")
                    try:
                        # Try direct navigation first if we're not already on the search page
                        if "/activity/search" not in page.url:
                            search_link = page.wait_for_selector("a[href='/activity/search']", timeout=5000)
                            if search_link:
                                search_link.click()
                                print("Clicked Custom Date Range link")
                                time.sleep(5)
                            else:
                                print("Custom Date Range link not found, trying direct navigation")
                                page.goto("https://www.americanexpress.com/en-us/account/activity/search", wait_until='networkidle', timeout=30000)
                        else:
                            print("Already on Custom Date Range page")
                    except Exception as e:
                        print(f"Error navigating to Custom Date Range: {e}")
                        # Try direct navigation
                        try:
                            page.goto("https://www.americanexpress.com/en-us/account/activity/search", wait_until='networkidle', timeout=30000)
                        except Exception as e:
                            print(f"Direct navigation to search page failed: {e}")
                    
                    time.sleep(5)
                    
                    # Take screenshot of the search page
                    page.screenshot(path=os.path.join(screenshots_dir, "search_page.png"))
                    
                    # Check if we got a "Page Not Found" error
                    if page.query_selector("text='Page Not Found'"):
                        print("Page Not Found error encountered. Trying alternative approach...")
                        
                        # Go back to the main account page
                        try:
                            page.click("text='Go back to the previous page'")
                            time.sleep(3)
                        except:
                            try:
                                page.click("text='Go to American Express Homepage'")
                                time.sleep(3)
                                
                                # If we went to homepage, we need to log back in
                                try:
                                    page.click("text='Log In'")
                                    time.sleep(3)
                                    page.fill("#eliloUserID", amex_username)
                                    page.fill("#eliloPassword", amex_password)
                                    page.click("#loginSubmit")
                                    time.sleep(10)
                                except Exception as e:
                                    print(f"Error logging back in: {e}")
                            except:
                                # As a last resort, go directly to the account home
                                page.goto("https://www.americanexpress.com/en-us/account/home", wait_until='networkidle', timeout=30000)
                                time.sleep(5)
                        
                        # Now try to go to statements page using a different approach
                        try:
                            print("Trying to go to Statements directly...")
                            page.goto("https://www.americanexpress.com/en-us/account/statements", wait_until='networkidle', timeout=30000)
                            time.sleep(5)
                            
                            # Take screenshot of where we landed
                            page.screenshot(path=os.path.join(screenshots_dir, "statements_direct_navigation.png"))
                            
                            # Look for a way to download transactions
                            download_links = page.query_selector_all("a:has-text('Download'), button:has-text('Download')")
                            if download_links and len(download_links) > 0:
                                print(f"Found {len(download_links)} download links")
                                download_links[0].click()
                                print("Clicked first download link")
                                time.sleep(5)
                            else:
                                print("No download links found on statements page")
                                # Try the main activity page instead
                                page.goto("https://www.americanexpress.com/en-us/account/activity", wait_until='networkidle', timeout=30000)
                                time.sleep(5)
                                
                                # Take screenshot of activity page
                                page.screenshot(path=os.path.join(screenshots_dir, "activity_page.png"))
                        except Exception as e:
                            print(f"Error with alternative navigation: {e}")
                    
                    # Click search button
                    print("\nClicking search button...")
                    try:
                        # First, check if there's a date range picker and set it if needed
                        try:
                            start_date = page.query_selector("input[placeholder='MM/DD/YYYY'], input[aria-label*='Start Date']")
                            end_date = page.query_selector("input[placeholder='MM/DD/YYYY']:nth-of-type(2), input[aria-label*='End Date']")
                            
                            if start_date and end_date:
                                print("Found date inputs, setting date range")
                                # Set to last 30 days
                                today = datetime.datetime.now()
                                thirty_days_ago = today - datetime.timedelta(days=30)
                                
                                start_date.fill(thirty_days_ago.strftime("%m/%d/%Y"))
                                end_date.fill(today.strftime("%m/%d/%Y"))
                                print("Date range set")
                        except Exception as e:
                            print(f"Error setting date range: {e}")
                        
                        # Now click search button
                        search_buttons = page.query_selector_all("button:has-text('Search'), [role='button']:has-text('Search')")
                        if len(search_buttons) >= 3:
                            print(f"Found {len(search_buttons)} search buttons. Clicking the 3rd one...")
                            search_buttons[2].click()
                        elif len(search_buttons) > 0:
                            print(f"Found {len(search_buttons)} search buttons. Clicking the last one...")
                            search_buttons[-1].click()
                        else:
                            print("No search buttons found, looking for a primary button")
                            primary_button = page.query_selector("button.btn-primary, button.axp-activity__cta--primary")
                            if primary_button:
                                primary_button.click()
                                print("Clicked primary button")
                            else:
                                print("No primary button found. Looking for any button in the search form")
                                form_buttons = page.query_selector_all("form button")
                                if form_buttons and len(form_buttons) > 0:
                                    form_buttons[-1].click()
                                    print("Clicked last form button")
                    except Exception as e:
                        print(f"Error clicking search button: {e}")
                    
                    # Wait for search results
                    time.sleep(10)
                    
                    # Take screenshot before clicking download button
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
                    
                    # Define download handlers to capture downloads
                    download_started = False
                    
                    def handle_download(download):
                        nonlocal download_started
                        print(f"\n*** Download started: {download.suggested_filename} ***")
                        # Create more informative filename with date and card name
                        current_date = datetime.datetime.now().strftime("%Y%m%d")
                        card_identifier = "Platinum_Card"
                        
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
                    
                    # Now handle the dialog for file type selection
                    print("\nHandling download dialog...")
                    
                    # Wait for dialog to appear
                    time.sleep(3)
                    
                    # Take screenshot of dialog for debugging
                    page.screenshot(path=os.path.join(screenshots_dir, "download_dialog.png"))
                    
                    # PLATINUM CARD SPECIFIC: Select Excel format if needed
                    try:
                        print("Checking if Excel option is already selected...")
                        excel_radio = page.query_selector("input[type='radio'][id*='excel']")
                        if excel_radio:
                            print("Found Excel radio button, checking if selected")
                            if not excel_radio.is_checked():
                                print("Excel not selected, clicking it")
                                excel_radio.click()
                                print("Clicked Excel radio button")
                            else:
                                print("Excel already selected")
                    except Exception as e:
                        print(f"Error with Excel selection: {e}")
                        # Try clicking the Excel option by text
                        try:
                            page.click("text='Excel'")
                            print("Clicked Excel option by text")
                        except Exception as e:
                            print(f"Failed to click Excel by text: {e}")
                    
                    time.sleep(2)
                    
                    # PLATINUM CARD SPECIFIC: Click the blue Download button
                    print("\nClicking the blue Download button...")
                    
                    # Method 1: Try direct selector for the blue button
                    try:
                        blue_button = page.query_selector("button.axp-activity__cta--primary")
                        if blue_button:
                            print("Found blue Download button by class")
                            blue_button.click()
                            print("Clicked blue Download button")
                        else:
                            print("Blue button not found by class")
                            
                            # Try by role and text
                            buttons = page.query_selector_all("[role='dialog'] button")
                            print(f"Found {len(buttons)} buttons in dialog")
                            
                            for i, button in enumerate(buttons):
                                try:
                                    text = button.inner_text()
                                    print(f"Button {i} text: {text}")
                                    if "download" in text.lower():
                                        print(f"Found Download button at index {i}")
                                        button.click()
                                        print("Clicked Download button")
                                        break
                                except Exception as e:
                                    print(f"Error checking button {i}: {e}")
                            
                            # If still not found, click the last button (usually the primary action)
                            if not download_started and len(buttons) > 0:
                                print("Clicking the last button in the dialog")
                                buttons[-1].click()
                                print("Clicked last button")
                    except Exception as e:
                        print(f"Error finding blue button: {e}")
                    
                    time.sleep(5)
                    
                    # Method 2: Use exact coordinates from the screenshot
                    if not download_started:
                        print("Using exact coordinates for Download button")
                        try:
                            # Click on the blue Download button in the bottom right of the dialog
                            # These coordinates are based on the screenshot
                            page.mouse.click(800, 564)
                            print("Clicked at coordinates (800, 564)")
                            time.sleep(5)
                        except Exception as e:
                            print(f"Error clicking at coordinates: {e}")
                    
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
                        print(f"- Card: Platinum Card")
                        print(f"- Download location: {download_dir}")
                        
                        # Write a summary to a log file
                        log_file = os.path.join(logs_dir, f"download_log_platinum_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
                        
                        try:
                            with open(log_file, 'w') as f:
                                f.write(f"Download Summary\n")
                                f.write(f"---------------\n")
                                f.write(f"Date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                                f.write(f"Card: Platinum Card\n")
                                f.write(f"Download location: {download_dir}\n")
                                f.write(f"Download successful: Yes\n")
                            print(f"Download log saved to: {log_file}")
                        except Exception as e:
                            print(f"Could not write log file: {e}")
                    else:
                        print("\nDownload did not start within the timeout period.")
                        
                        # Write a failure log
                        log_file = os.path.join(logs_dir, f"download_log_platinum_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_FAILED.txt")
                        
                        try:
                            with open(log_file, 'w') as f:
                                f.write(f"Download Summary\n")
                                f.write(f"---------------\n")
                                f.write(f"Date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                                f.write(f"Card: Platinum Card\n")
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
                                if os.path.getmtime(file_path) > recent_time and (file.endswith('.xlsx') or file.endswith('.xls') or 'amex' in file.lower() or 'platinum' in file.lower()):
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
                try:
                    page.screenshot(path=os.path.join(screenshots_dir, f"error_screenshot_platinum_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.png"))
                    print(f"Error screenshot saved")
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
    main()