# AMEX Statement Downloader

Automated script to download American Express statements using Playwright browser automation.

## Features

- Automated login with 2FA support
- Downloads transaction data in Excel format
- Handles AMEX Gold Card statements
- Custom date range selection
- Automatic file organization

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Create config/.env file with your credentials:
```
AMEX_USERNAME=your_username
AMEX_PASSWORD=your_password
```

3. Set up OTP file path for 2FA (currently configured for iCloud Drive)

## Usage

```bash
python amex_gold_downloader.py
```

The script will:
1. Open American Express login page
2. Handle 2FA authentication
3. Navigate to statements section
4. Download transaction data to ~/Downloads/AmexStatements/

## Output

Downloaded files are saved to the `output/` directory and organized by date.