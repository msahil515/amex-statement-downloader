#!/usr/bin/env python
"""
Script to download statements for multiple American Express cards.
"""
import argparse
import subprocess
import time
import os
from datetime import datetime

# List of commonly used Amex card names
DEFAULT_CARD_NAMES = [
    "American Express Gold Card", 
    "Platinum Card",
    "Blue Business Plus Card",
    "Green Card",
    "Hilton Honors Aspire Card",
    "Marriott Bonvoy Brilliant American Express Card",
    "Business Gold Card"
]

def download_for_card(card_name, wait_time=60):
    """
    Run the downloader script for a specific card.
    
    Args:
        card_name (str): The name of the card to download statements for
        wait_time (int): Time to wait in seconds between card runs
    """
    print(f"\n{'='*50}")
    print(f"Starting download for: {card_name}")
    print(f"{'='*50}")
    
    # Log start time
    start_time = datetime.now()
    print(f"Start time: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Run the downloader script with the specified card
    cmd = ["python", "amex_gold_downloader_modified.py", "--card", card_name]
    print(f"Running command: {' '.join(cmd)}")
    
    try:
        subprocess.run(cmd, timeout=600)  # 10-minute timeout
        print(f"\nDownload process for {card_name} completed!")
    except subprocess.TimeoutExpired:
        print(f"\nWarning: Download process for {card_name} timed out after 10 minutes.")
        print("This might be normal if 2FA verification is required. Please check the browser window.")
    except Exception as e:
        print(f"\nError running download for {card_name}: {e}")
    
    # Log end time
    end_time = datetime.now()
    duration = end_time - start_time
    print(f"End time: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Duration: {duration}")
    print(f"{'='*50}\n")
    
    # Wait between runs
    if wait_time > 0:
        print(f"Waiting {wait_time} seconds before proceeding to next card...")
        time.sleep(wait_time)

def main():
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Download statements for multiple Amex cards')
    parser.add_argument('--cards', nargs='+', help='List of card names to download statements for')
    parser.add_argument('--wait', type=int, default=60, help='Wait time in seconds between card downloads')
    
    args = parser.parse_args()
    
    # Use provided cards or default list
    cards_to_process = args.cards if args.cards else DEFAULT_CARD_NAMES
    
    # Create logs directory if it doesn't exist
    log_dir = os.path.join(os.path.dirname(__file__), 'logs')
    os.makedirs(log_dir, exist_ok=True)
    
    # Log the run
    log_file = os.path.join(log_dir, f"multi_card_run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
    with open(log_file, 'w') as f:
        f.write(f"Multi-Card Download Run\n")
        f.write(f"======================\n")
        f.write(f"Start time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Cards to process:\n")
        for card in cards_to_process:
            f.write(f"- {card}\n")
        f.write(f"Wait time between cards: {args.wait} seconds\n\n")
    
    print(f"Starting multi-card download process")
    print(f"Log file: {log_file}")
    print(f"Processing {len(cards_to_process)} cards: {', '.join(cards_to_process)}")
    
    # Process each card
    for i, card in enumerate(cards_to_process, 1):
        print(f"\nProcessing card {i}/{len(cards_to_process)}")
        download_for_card(card, args.wait)
    
    # Update log with completion
    with open(log_file, 'a') as f:
        f.write(f"\nProcess completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    print("\nMulti-card download process completed!")
    print(f"Check {log_file} for details.")

if __name__ == "__main__":
    main()