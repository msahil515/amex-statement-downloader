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
            print("Pausing execution - we'll implement the next steps in subsequent updates.")
            
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