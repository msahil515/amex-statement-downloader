#!/usr/bin/env python
"""
Basic script for downloading American Express statements.
This is a simplified version with less advanced error handling.
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
                
                # Make it clear we need manual intervention
                print("\n==============================================================")
                print("| MANUAL ACTION REQUIRED                                     |")
                print("==============================================================")
                print("| The browser is now at the dashboard.                       |")
                print("| Please verify you're logged in.                            |")
                print("|                                                            |")
                print("| IMPORTANT: Continue script execution in the terminal with  |")
                print("| the 'c' command (not Enter).                              |")
                print("==============================================================")
                
                # Use a timeout approach instead of blocking on input
                # This way the script will continue after a delay
                continue_execution = False
                for i in range(30):  # Wait up to 30 seconds
                    print(f"Continuing in {30-i} seconds... (type 'c' to continue now)")
                    try:
                        # Use non-blocking input with a timeout
                        import select
                        import sys
                        i, o, e = select.select([sys.stdin], [], [], 1)
                        if i and sys.stdin.readline().strip().lower() == 'c':
                            print("Continuing now...")
                            continue_execution = True
                            break
                    except:
                        # If select is not supported, just sleep
                        time.sleep(1)
                
                print("Continuing with script execution...")
                
                # Take a screenshot to see where we are
                screenshots_dir = os.path.join(os.path.dirname(__file__), 'screenshots')
                os.makedirs(screenshots_dir, exist_ok=True)
                page.screenshot(path=os.path.join(screenshots_dir, "after_login_dashboard.png"))
                
                # Select card based on parameter or default to Gold Card
                card_to_select = card_name or "American Express Gold Card"
                print(f"\nSelecting card: {card_to_select}...")
                
                # Click on card selector dropdown
                try:
                    page.click("[role='combobox']")
                    time.sleep(3)
                    print("Clicked on card dropdown")
                    
                    # Wait for the user to select the card manually
                    print("\n==============================================================")
                    print("| MANUAL ACTION REQUIRED                                     |")
                    print("==============================================================")
                    print(f"| Please manually select {card_to_select} in the dropdown.    |")
                    print("|                                                            |")
                    print("| Type 'c' when done or wait 30 seconds for auto-continue    |")
                    print("==============================================================")
                    
                    # Auto-continue after timeout
                    for i in range(30):  # Wait up to 30 seconds
                        print(f"Continuing in {30-i} seconds... (type 'c' to continue now)")
                        try:
                            # Use non-blocking input with a timeout
                            import select
                            i, o, e = select.select([sys.stdin], [], [], 1)
                            if i and sys.stdin.readline().strip().lower() == 'c':
                                print("Continuing now...")
                                break
                        except:
                            # If select is not supported, just sleep
                            time.sleep(1)
                    print("Continuing script execution...")
                except Exception as e:
                    print(f"Error clicking card dropdown: {e}")
                    print("If you're already on the correct card, that's fine.")
                
                # Now navigate to Statements & Activity
                print("\nNavigating to Statements & Activity...")
                try:
                    page.click("span:has-text('Statements & Activity')")
                    print("Clicked on Statements & Activity")
                    time.sleep(5)
                except Exception as e:
                    print(f"Error clicking Statements & Activity: {e}")
                    print("\n==============================================================")
                    print("| MANUAL ACTION REQUIRED                                     |")
                    print("==============================================================")
                    print("| Please navigate to Statements & Activity manually.         |")
                    print("|                                                            |")
                    print("| Type 'c' when done or wait 30 seconds for auto-continue    |")
                    print("==============================================================")
                    
                    # Auto-continue after timeout
                    for i in range(30):  # Wait up to 30 seconds
                        print(f"Continuing in {30-i} seconds... (type 'c' to continue now)")
                        try:
                            # Use non-blocking input with a timeout
                            import select
                            i, o, e = select.select([sys.stdin], [], [], 1)
                            if i and sys.stdin.readline().strip().lower() == 'c':
                                print("Continuing now...")
                                break
                        except:
                            # If select is not supported, just sleep
                            time.sleep(1)
                    print("Continuing script execution...")
                
                # Take a screenshot to see where we are
                page.screenshot(path=os.path.join(screenshots_dir, "statements_activity_page.png"))
                
                # Go to Custom Date Range
                print("\nNavigating to Custom Date Range...")
                try:
                    page.click("a[href='/activity/search']")
                    print("Clicked on Custom Date Range")
                    time.sleep(5)
                except Exception as e:
                    print(f"Error navigating to Custom Date Range: {e}")
                    print("\n==============================================================")
                    print("| MANUAL ACTION REQUIRED                                     |")
                    print("==============================================================")
                    print("| Please navigate to Custom Date Range manually.             |")
                    print("|                                                            |")
                    print("| Type 'c' when done or wait 30 seconds for auto-continue    |")
                    print("==============================================================")
                    
                    # Auto-continue after timeout
                    for i in range(30):  # Wait up to 30 seconds
                        print(f"Continuing in {30-i} seconds... (type 'c' to continue now)")
                        try:
                            # Use non-blocking input with a timeout
                            import select
                            i, o, e = select.select([sys.stdin], [], [], 1)
                            if i and sys.stdin.readline().strip().lower() == 'c':
                                print("Continuing now...")
                                break
                        except:
                            # If select is not supported, just sleep
                            time.sleep(1)
                    print("Continuing script execution...")
                
                # Take a screenshot to see where we are
                page.screenshot(path=os.path.join(screenshots_dir, "custom_date_range_page.png"))
                
                # Wait for user to confirm
                print("\n==============================================================")
                print("| MANUAL ACTION REQUIRED                                     |")
                print("==============================================================")
                print("| Please verify that you're on the custom date range search page.|")
                print("|                                                            |")
                print("| Type 'c' when ready or wait 30 seconds for auto-continue    |")
                print("==============================================================")
                
                # Auto-continue after timeout
                for i in range(30):  # Wait up to 30 seconds
                    print(f"Continuing in {30-i} seconds... (type 'c' to continue now)")
                    try:
                        # Use non-blocking input with a timeout
                        import select
                        i, o, e = select.select([sys.stdin], [], [], 1)
                        if i and sys.stdin.readline().strip().lower() == 'c':
                            print("Continuing now...")
                            break
                    except:
                        # If select is not supported, just sleep
                        time.sleep(1)
                print("Continuing script execution...")
                
                # Click search button
                print("\nClicking search button...")
                try:
                    search_buttons = page.query_selector_all("button:has-text('Search'), [role='button']:has-text('Search')")
                    if len(search_buttons) >= 3:
                        print(f"Found {len(search_buttons)} search buttons. Clicking the 3rd one...")
                        search_buttons[2].click()
                    else:
                        print(f"Not enough search buttons found (found {len(search_buttons)}, need at least 3)")
                        # Fallback to clicking the last one
                        if search_buttons:
                            print("Clicking last available search button")
                            search_buttons[-1].click()
                        else:
                            print("No search buttons found")
                            print("\n==============================================================")
                            print("| MANUAL ACTION REQUIRED                                     |")
                            print("==============================================================")
                            print("| Please click the Search button manually.                   |")
                            print("|                                                            |")
                            print("| Type 'c' when done or wait 30 seconds for auto-continue    |")
                            print("==============================================================")
                            
                            # Auto-continue after timeout
                            for i in range(30):  # Wait up to 30 seconds
                                print(f"Continuing in {30-i} seconds... (type 'c' to continue now)")
                                try:
                                    # Use non-blocking input with a timeout
                                    import select
                                    i, o, e = select.select([sys.stdin], [], [], 1)
                                    if i and sys.stdin.readline().strip().lower() == 'c':
                                        print("Continuing now...")
                                        break
                                except:
                                    # If select is not supported, just sleep
                                    time.sleep(1)
                            print("Continuing script execution...")
                except Exception as e:
                    print(f"Error clicking search button: {e}")
                    print("\n==============================================================")
                    print("| MANUAL ACTION REQUIRED                                     |")
                    print("==============================================================")
                    print("| Please click the Search button manually.                   |")
                    print("|                                                            |")
                    print("| Type 'c' when done or wait 30 seconds for auto-continue    |")
                    print("==============================================================")
                    
                    # Auto-continue after timeout
                    for i in range(30):  # Wait up to 30 seconds
                        print(f"Continuing in {30-i} seconds... (type 'c' to continue now)")
                        try:
                            # Use non-blocking input with a timeout
                            import select
                            i, o, e = select.select([sys.stdin], [], [], 1)
                            if i and sys.stdin.readline().strip().lower() == 'c':
                                print("Continuing now...")
                                break
                        except:
                            # If select is not supported, just sleep
                            time.sleep(1)
                    print("Continuing script execution...")
                
                # Wait for search results
                print("\nWaiting for search results...")
                time.sleep(10)
                
                # Take a screenshot of search results
                page.screenshot(path=os.path.join(screenshots_dir, "search_results.png"))
                
                # Click download button to open dialog
                print("\nClicking download button to open dialog...")
                try:
                    download_button = page.wait_for_selector("button:has-text('Download')")
                    download_button.click()
                    print("Clicked download button")
                except Exception as e:
                    print(f"Error finding download button: {e}")
                    print("\n==============================================================")
                    print("| MANUAL ACTION REQUIRED                                     |")
                    print("==============================================================")
                    print("| Please click the Download button manually.                  |")
                    print("|                                                            |")
                    print("| Type 'c' when done or wait 30 seconds for auto-continue    |")
                    print("==============================================================")
                    
                    # Auto-continue after timeout
                    for i in range(30):  # Wait up to 30 seconds
                        print(f"Continuing in {30-i} seconds... (type 'c' to continue now)")
                        try:
                            # Use non-blocking input with a timeout
                            import select
                            i, o, e = select.select([sys.stdin], [], [], 1)
                            if i and sys.stdin.readline().strip().lower() == 'c':
                                print("Continuing now...")
                                break
                        except:
                            # If select is not supported, just sleep
                            time.sleep(1)
                    print("Continuing script execution...")
                
                time.sleep(3)
                
                # Take a screenshot of the dialog
                page.screenshot(path=os.path.join(screenshots_dir, "download_dialog.png"))
                
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
                
                # Wait for user to manually click the Download button in the dialog
                print("\n==============================================================")
                print("| MANUAL ACTION REQUIRED                                     |")
                print("==============================================================")
                print("| Please click the blue Download button in the dialog.        |")
                print("|                                                            |")
                print("| Type 'c' when done or wait 30 seconds for auto-continue    |")
                print("==============================================================")
                
                # Auto-continue after timeout
                for i in range(30):  # Wait up to 30 seconds
                    print(f"Continuing in {30-i} seconds... (type 'c' to continue now)")
                    try:
                        # Use non-blocking input with a timeout
                        import select
                        i, o, e = select.select([sys.stdin], [], [], 1)
                        if i and sys.stdin.readline().strip().lower() == 'c':
                            print("Continuing now...")
                            break
                    except:
                        # If select is not supported, just sleep
                        time.sleep(1)
                print("Continuing script execution...")
                
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
                    print(f"- Card: {card_name or 'Default Card'}")
                    print(f"- Download location: {download_dir}")
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
            screenshots_dir = os.path.join(os.path.dirname(__file__), 'screenshots')
            os.makedirs(screenshots_dir, exist_ok=True)
            try:
                page.screenshot(path=os.path.join(screenshots_dir, "error_screenshot.png"))
                print(f"Error screenshot saved to {os.path.join(screenshots_dir, 'error_screenshot.png')}")
            except:
                print("Could not take error screenshot")
        
        finally:
            # Ask user if they want to keep the browser open
            try:
                print("\n==============================================================")
                print("| SCRIPT COMPLETED                                           |")
                print("==============================================================")
                print("| Do you want to keep the browser open?                      |")
                print("|                                                            |")
                print("| Type 'y' to keep open, or browser will close in 10 seconds |")
                print("==============================================================")
                
                keep_open = False
                for i in range(10):  # Wait up to 10 seconds
                    print(f"Closing browser in {10-i} seconds... (type 'y' to keep open)")
                    try:
                        import select
                        i, o, e = select.select([sys.stdin], [], [], 1)
                        if i and sys.stdin.readline().strip().lower() == 'y':
                            keep_open = True
                            print("Browser will remain open.")
                            break
                    except:
                        time.sleep(1)
                if keep_open != 'y':
                    # Close browser
                    print("Closing browser...")
                    browser.close()
                else:
                    print("Browser left open. You'll need to close it manually.")
            except:
                # Close browser
                print("\nClosing browser...")
                browser.close()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Download American Express statements with manual assistance')
    parser.add_argument('--card', type=str, help='Card name to select (e.g., "American Express Gold Card", "Platinum Card")')
    args = parser.parse_args()
    
    main(card_name=args.card)