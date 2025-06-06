#!/usr/bin/env python
"""
Focused script to click the Download button in the AMEX dialog.
This script assumes a browser is already open at the download dialog.
"""
from playwright.sync_api import sync_playwright
import time
import os
import sys

def click_download_button():
    """Try multiple approaches to click the Download button in the dialog."""
    screenshots_dir = os.path.join(os.path.dirname(__file__), 'screenshots')
    os.makedirs(screenshots_dir, exist_ok=True)
    
    with sync_playwright() as p:
        # Use a new context with a view of the existing page
        browser = p.chromium.launch(
            headless=False,
            args=[
                '--disable-blink-features=AutomationControlled',
                '--disable-features=IsolateOrigins,site-per-process',
                '--disable-site-isolation-trials'
            ]
        )
        
        context = browser.new_context(
            viewport={'width': 1920, 'height': 1080},
            user_agent='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
        )
        
        page = context.new_page()
        
        try:
            # Go directly to American Express activity page
            page.goto("https://www.americanexpress.com/en-us/account/activity", wait_until='networkidle')
            time.sleep(5)
            
            # Take a screenshot of where we are
            page.screenshot(path=os.path.join(screenshots_dir, "activity_page_start.png"))
            
            print("1. FIRST, please navigate to your search results manually.")
            print("2. Click the first 'Download' button to open the dialog box.")
            print("3. When the dialog box is open, type 'ready' below:")
            
            ready = input("Type 'ready' when the dialog is showing: ")
            if ready.lower() != 'ready':
                print("Please type 'ready' when the dialog is showing.")
                return
            
            # Take a screenshot of the dialog
            page.screenshot(path=os.path.join(screenshots_dir, "dialog_before_click.png"))
            
            print("\nTrying multiple approaches to click the Download button...")
            
            # APPROACH 1: JavaScript click on the blue button
            print("\nApproach 1: Using JavaScript to find and click the blue Download button")
            try:
                # This JavaScript finds and clicks the blue Download button based on text content and background color
                result = page.evaluate("""
                    (() => {
                        // Get all elements with text 'Download'
                        const elements = Array.from(document.querySelectorAll('*')).filter(el => 
                            el.textContent.trim() === 'Download');
                        
                        console.log('Found ' + elements.length + ' elements with text Download');
                        
                        // Try to find the blue button
                        for (const el of elements) {
                            const style = window.getComputedStyle(el);
                            const bg = style.backgroundColor;
                            const isBlue = bg.includes('0, 111, 207') || bg.includes('rgb(0, 0, 255)') || 
                                         bg.includes('rgb(0, 11') || bg.includes('rgb(0, 10') || 
                                         bg.includes('rgba(0, 1');
                            
                            if (isBlue || el.classList.contains('btn-primary') || 
                                el.closest('button')?.classList.contains('btn-primary')) {
                                console.log('Found blue Download button');
                                // Click the element or its parent button
                                const buttonToClick = el.tagName === 'BUTTON' ? el : 
                                                    el.closest('button') || el;
                                buttonToClick.click();
                                return true;
                            }
                        }
                        
                        // If no blue button with "Download" text, try to find any blue button in the dialog
                        const dialog = document.querySelector('[role="dialog"]');
                        if (dialog) {
                            const buttons = Array.from(dialog.querySelectorAll('button'));
                            for (const button of buttons) {
                                const style = window.getComputedStyle(button);
                                const bg = style.backgroundColor;
                                const isBlue = bg.includes('0, 111, 207') || bg.includes('rgb(0, 0, 255)') || 
                                             bg.includes('rgb(0, 11') || bg.includes('rgb(0, 10') || 
                                             bg.includes('rgba(0, 1');
                                
                                if (isBlue) {
                                    console.log('Found blue button in dialog');
                                    button.click();
                                    return true;
                                }
                            }
                            
                            // If no blue button found, click the last button (typically the primary action)
                            if (buttons.length > 0) {
                                console.log('Clicking last button in dialog');
                                buttons[buttons.length - 1].click();
                                return true;
                            }
                        }
                        
                        return false;
                    })();
                """)
                
                print(f"JavaScript approach result: {result}")
                if result:
                    print("JavaScript click seems to have worked")
                    time.sleep(5)
            except Exception as e:
                print(f"JavaScript approach failed: {e}")
            
            # Take a screenshot after JavaScript approach
            page.screenshot(path=os.path.join(screenshots_dir, "after_javascript_click.png"))
            
            # APPROACH 2: Try direct selectors
            print("\nApproach 2: Using direct selectors")
            selectors = [
                "[role='dialog'] button.axp-activity__cta--primary",
                "[role='dialog'] button.btn-primary",
                "[role='dialog'] button:has-text('Download')",
                "[role='dialog'] button.cta-primary",
                "[role='dialog'] button:last-child",
                ".modal-footer button:last-child",
                ".modal-footer button.btn-primary",
                ".download-button"
            ]
            
            for selector in selectors:
                try:
                    print(f"Trying selector: {selector}")
                    elements = page.query_selector_all(selector)
                    print(f"Found {len(elements)} elements matching {selector}")
                    
                    if elements:
                        # Click the last element (usually the primary action)
                        elements[-1].click()
                        print(f"Clicked element using selector: {selector}")
                        time.sleep(5)
                        # Take a screenshot after click
                        page.screenshot(path=os.path.join(screenshots_dir, f"after_click_{selector.replace(':', '_').replace('[', '').replace(']', '')}.png"))
                except Exception as e:
                    print(f"Error with selector {selector}: {e}")
            
            # APPROACH 3: Direct coordinates based on dialog position
            print("\nApproach 3: Using direct coordinates at likely locations")
            
            # Get dialog dimensions if possible
            try:
                dialog_box = page.query_selector("[role='dialog']")
                if dialog_box:
                    box = dialog_box.bounding_box()
                    if box:
                        print(f"Dialog box position: x={box['x']}, y={box['y']}, width={box['width']}, height={box['height']}")
                        
                        # Click in the bottom right corner of the dialog (where action buttons typically are)
                        bottom_right_x = box['x'] + box['width'] - 80  # 80px from right edge
                        bottom_right_y = box['y'] + box['height'] - 30  # 30px from bottom edge
                        
                        print(f"Clicking at bottom right: ({bottom_right_x}, {bottom_right_y})")
                        page.mouse.click(bottom_right_x, bottom_right_y)
                        time.sleep(5)
                        
                        # Take screenshot after click
                        page.screenshot(path=os.path.join(screenshots_dir, "after_bottom_right_click.png"))
                else:
                    print("Could not find dialog element")
            except Exception as e:
                print(f"Error with coordinate approach: {e}")
            
            # APPROACH 4: Grid of clicks across the bottom of the screen
            print("\nApproach 4: Grid of clicks across bottom of screen")
            
            # Create a grid of points along the bottom of the viewport where action buttons often are
            width = page.viewport_size['width']
            height = page.viewport_size['height']
            
            # Focus on the bottom right quadrant
            start_x = width // 2
            end_x = width - 50
            start_y = height - 200
            end_y = height - 50
            
            grid_points = []
            for x in range(start_x, end_x, 100):  # Every 100 pixels horizontally
                for y in range(start_y, end_y, 50):  # Every 50 pixels vertically
                    grid_points.append((x, y))
            
            # Add specific points where the Download button is likely to be
            # These are the most likely positions based on common dialog layouts
            grid_points = [
                (800, 564),  # Common position based on screenshot
                (750, 564),  # Slightly to the left
                (850, 564),  # Slightly to the right
                (800, 540),  # Slightly higher
                (800, 590),  # Slightly lower
            ] + grid_points
            
            for i, (x, y) in enumerate(grid_points):
                try:
                    print(f"Clicking at grid point {i+1}/{len(grid_points)}: ({x}, {y})")
                    page.mouse.click(x, y)
                    time.sleep(3)
                    
                    # Take screenshot after every 5 clicks
                    if (i + 1) % 5 == 0:
                        page.screenshot(path=os.path.join(screenshots_dir, f"after_grid_clicks_{i+1}.png"))
                except Exception as e:
                    print(f"Error clicking at ({x}, {y}): {e}")
            
            print("\nAll approaches tried. Please check if the download started.")
            print("Look in the following locations for downloaded files:")
            print("1. ~/Downloads/AmexStatements/Platinum_Card/")
            print("2. ~/Downloads/ (for any recent Excel files)")
            
            # Keep browser open
            print("\nBrowser will remain open so you can manually complete the download.")
            input("Press Enter when you're done to close the browser...")
            
        except Exception as e:
            print(f"Error: {e}")
            # Take error screenshot
            page.screenshot(path=os.path.join(screenshots_dir, "click_download_error.png"))
        
        finally:
            browser.close()

if __name__ == "__main__":
    click_download_button()