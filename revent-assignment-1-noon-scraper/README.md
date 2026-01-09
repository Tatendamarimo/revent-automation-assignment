# Noon Web Scraper

A Python web scraper for noon.com that extracts product listings with multi-seller support and exports to Excel.

## Features

- Search for products by keyword
- Automatically detects multiple sellers for each product
- Extracts 9 data fields per product-seller combination
- Exports to formatted Excel files
- Includes delays to avoid bot detection
- Real-time progress tracking

## Requirements

- Python 3.7+
- Chrome browser
- Internet connection

## Installation

Clone the repository and install dependencies:

```bash
pip install -r requirements.txt
```

This installs:
- `selenium` - Web automation
- `webdriver-manager` - Automatic ChromeDriver management
- `pandas` - Data manipulation
- `openpyxl` - Excel file creation

## Usage

Run the scraper:

```bash
python main.py
```

Follow the prompts:
1. Enter search keywords (one per line, type "done" when finished)
2. Optionally set a limit for products per keyword
3. Confirm to start scraping

Example:
```
Keyword 1: iphone
Keyword 2: samsung galaxy
Keyword 3: done

Limit products per keyword? (Enter number or press Enter for all): 10

Proceed with scraping? (yes/no): yes
```

## Output

The scraper creates an Excel file in the `output/` directory:

| Search Keyword | Category | Title | Description | Price | Rating | Reviews | Seller | Product URL |
|----------------|----------|-------|-------------|-------|--------|---------|--------|-------------|
| iphone | Electronics > Mobiles | iPhone 16... | 128GB, White... | AED 3,999 | 4.6 | 12.8K | noon | https://... |

Note: Products with multiple sellers will have one row per seller.

Files are named with timestamps:
- Single keyword: `noon_scraper_iphone_2026-01-09_13-30-45.xlsx`
- Multiple keywords: `noon_scraper_2026-01-09_13-30-45.xlsx`

## Data Fields

1. Search Keyword - The keyword used to find the product
2. Category - Product category from breadcrumbs
3. Title - Full product title
4. Description - Product highlights/description
5. Price - Product price (varies by seller)
6. Rating - Product rating (1-5 stars)
7. Reviews - Number of reviews
8. Seller - Seller name (e.g., "noon", "TechStore")
9. Product URL - Direct link to product page

## Configuration

You can edit `config.py` to customize:
- Timeouts and delays
- CSS selectors (if website structure changes)
- Output directory and filename format

## Troubleshooting

**ChromeDriver not found**
- The script automatically downloads ChromeDriver
- Make sure Chrome browser is installed

**No data scraped**
- Check your internet connection
- Verify noon.com is accessible
- Check `scraper.log` for error details

**CAPTCHA appears**
- The scraper includes delays to avoid this
- If it happens, you may need to solve it manually
- Try increasing delays in `config.py`

**Selectors not working**
- noon.com may have updated their website
- Update CSS selectors in `config.py`
- Check `scraper.log` for specific errors

Check `scraper.log` for detailed execution logs.

## Project Structure

```
revent-assignment-1-noon-scraper/
├── main.py                 # Main entry point
├── noon_scraper.py         # Core scraper logic
├── excel_exporter.py       # Excel export functionality
├── config.py               # Configuration settings
├── test_scraper.py         # Test file
├── requirements.txt        # Python dependencies
├── README.md              # This file
├── output/                # Excel output directory (auto-created)
└── scraper.log           # Log file (auto-created)
```

## How It Works

- Uses Selenium WebDriver to handle JavaScript-rendered content
- Waits for elements to load before extraction
- Clicks "Other Sellers" button to access seller modal
- Continues scraping even if individual products fail
- Includes random delays (2-4 seconds) between requests
- Uses realistic user agent and disables automation flags

## Limitations

- Scraping is intentionally slow to avoid detection
- Large scrapes (100+ products) take time
- noon.com may implement additional anti-scraping measures
- Some fields may show "N/A" if not available
