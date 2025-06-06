#!/usr/bin/env python
"""
Simple utility to click the blue Download button in AMEX dialog.
"""
import pyautogui
import time
import os

def find_blue_button():
    # Create screenshots directory
    screenshots_dir = os.path.join(os.path.dirname(__file__), 'screenshots')
    os.makedirs(screenshots_dir, exist_ok=True)
    
    print("This utility will help you click the blue Download button in the AMEX dialog.")
    print("Please make sure the AMEX download dialog is visible on screen.")
    
    # Give user time to position the dialog
    print("\nYou have 5 seconds to make sure the dialog is visible...")
    for i in range(5, 0, -1):
        print(f"{i}...")
        time.sleep(1)
    
    print("\nTaking screenshot of current screen...")
    # Take a screenshot of the entire screen
    screen = pyautogui.screenshot()
    screen_path = os.path.join(screenshots_dir, "full_screen.png")
    screen.save(screen_path)
    print(f"Screenshot saved to {screen_path}")
    
    print("\nAnalyzing screenshot for blue buttons...")
    
    # Define the color range for the blue button (RGB)
    blue_min = (0, 90, 180)  # Minimum RGB values for blue
    blue_max = (60, 140, 255)  # Maximum RGB values for blue
    
    # Get screen dimensions
    width, height = screen.size
    
    # Initialize variables to store blue areas
    blue_areas = []
    
    # Scan the screen for blue pixels and identify blue areas
    for y in range(0, height, 10):  # Skip every 10 pixels for speed
        for x in range(0, width, 10):  # Skip every 10 pixels for speed
            pixel = screen.getpixel((x, y))
            
            # Check if pixel is in the blue range
            if (blue_min[0] <= pixel[0] <= blue_max[0] and
                blue_min[1] <= pixel[1] <= blue_max[1] and
                blue_min[2] <= pixel[2] <= blue_max[2]):
                
                # Found a blue pixel, check surrounding area
                area_size = 0
                for dy in range(max(0, y-20), min(height, y+20)):
                    for dx in range(max(0, x-20), min(width, x+20)):
                        try:
                            p = screen.getpixel((dx, dy))
                            if (blue_min[0] <= p[0] <= blue_max[0] and
                                blue_min[1] <= p[1] <= blue_max[1] and
                                blue_min[2] <= p[2] <= blue_max[2]):
                                area_size += 1
                        except:
                            pass
                
                # If we found a substantial blue area, add it to our list
                if area_size > 50:  # Arbitrary threshold
                    blue_areas.append((x, y, area_size))
    
    print(f"Found {len(blue_areas)} potential blue button areas")
    
    if blue_areas:
        # Sort by area size (largest first)
        blue_areas.sort(key=lambda a: a[2], reverse=True)
        
        print("\nTop 5 potential button locations:")
        for i, (x, y, size) in enumerate(blue_areas[:5]):
            print(f"{i+1}. Position: ({x}, {y}), Size: {size}")
        
        print("\nWill attempt to click the buttons in order from largest to smallest.")
        print("After each click, check if the download started.")
        
        for i, (x, y, size) in enumerate(blue_areas[:5]):
            input(f"\nPress Enter to click potential button {i+1} at ({x}, {y})...")
            
            # Move to and click the position
            pyautogui.moveTo(x, y, duration=1)
            pyautogui.click()
            
            print(f"Clicked at ({x}, {y})")
            
            # Ask if the download started
            result = input("Did the download start? (y/n): ")
            if result.lower() == 'y':
                print("Great! Download started successfully.")
                return
    else:
        print("No blue areas found. Trying common button locations.")
    
    # If no blue areas found or none worked, try common positions
    common_positions = [
        (800, 564),  # Common position based on previous runs
        (700, 564),  # Slightly to the left
        (900, 564),  # Slightly to the right
        (650, 450),  # Different area on screen
        (width//2, height//2 + 100),  # Below center of screen
        (width - 100, height - 100)  # Bottom right corner
    ]
    
    print("\nTrying common button positions:")
    for i, (x, y) in enumerate(common_positions):
        input(f"\nPress Enter to click common position {i+1} at ({x}, {y})...")
        
        # Move to and click the position
        pyautogui.moveTo(x, y, duration=1)
        pyautogui.click()
        
        print(f"Clicked at ({x}, {y})")
        
        # Ask if the download started
        result = input("Did the download start? (y/n): ")
        if result.lower() == 'y':
            print("Great! Download started successfully.")
            return
    
    print("\nCould not find the download button automatically.")
    print("As a last resort, you can use manual mode:")
    
    print("\nMANUAL MODE:")
    print("1. Move your mouse over the blue Download button")
    print("2. Note the position in the top-left corner of the utility")
    print("3. Enter the coordinates below")
    
    try:
        x = int(input("Enter X coordinate: "))
        y = int(input("Enter Y coordinate: "))
        
        # Move to and click the position
        pyautogui.moveTo(x, y, duration=1)
        pyautogui.click()
        
        print(f"Clicked at ({x}, {y})")
        
        # Ask if the download started
        result = input("Did the download start? (y/n): ")
        if result.lower() == 'y':
            print("Great! Download started successfully.")
            return
    except:
        print("Invalid coordinates")
    
    print("\nCould not complete the download automatically.")
    print("Please click the Download button manually.")

if __name__ == "__main__":
    # Check if pyautogui is installed
    try:
        import pyautogui
        
        # Pause between actions
        pyautogui.PAUSE = 1
        
        # Run the main function
        find_blue_button()
    except ImportError:
        print("Error: This utility requires the pyautogui package.")
        print("Please install it with: pip install pyautogui")
        print("Then run this utility again.")