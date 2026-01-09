# Revent Automation Assignment

This repository contains two Python automation assignments for web scraping and data processing.

## Project Structure

```
revent-automation-assignment/
├── revent-assignment-1-noon-scraper/    # Assignment 1: Noon Web Scraper
│   ├── main.py                          # Main entry point
│   ├── noon_scraper.py                  # Core scraper logic
│   ├── excel_exporter.py                # Excel export functionality
│   ├── config.py                        # Configuration settings
│   ├── test_scraper.py                  # Test file
│   ├── requirements.txt                 # Python dependencies
│   ├── README.md                        # Detailed documentation
│   └── output/                          # Generated Excel files
│
├── revent-assignment-2-report-merger/   # Assignment 2: Report Merger
│   ├── report_merger.py                 # Main script
│   ├── requirements.txt                 # Python dependencies
│   ├── README.md                        # Detailed documentation
│   ├── Assignments- Sheet Automation.xlsx           # Sample input
│   └── Assignments- Sheet Automation_MERGED.xlsx    # Sample output
│
└── README.md                            # This file
```

## Assignments Overview

### Assignment 1: Noon Web Scraper

A robust web scraper for noon.com that extracts product listings with multi-seller support.

**Key Features:**
- Search products by keyword
- Multi-seller detection and extraction
- 9 data fields per product-seller combination
- Automated Excel export with formatting
- Anti-bot detection measures

**Technologies Used:**
- `selenium` - Web automation and browser control
- `webdriver-manager` - Automatic ChromeDriver management
- `pandas` - Data manipulation and processing
- `openpyxl` - Excel file creation and formatting

[📖 Full Documentation](./revent-assignment-1-noon-scraper/README.md)

---

### Assignment 2: Amazon & Noon Report Merger

Automated script to merge Amazon and Noon sales reports using dynamic column mapping.

**Key Features:**
- Zero hardcoding - fully reusable
- Dynamic column mapping from configuration sheet
- Automatic data transformations
- Date component extraction
- Value calculations and conditional logic

**Technologies Used:**
- `pandas` - Data manipulation and merging
- `openpyxl` - Excel file reading and writing

[📖 Full Documentation](./revent-assignment-2-report-merger/README.md)

---

## Why This Solution?

### Technical Decisions

#### Assignment 1: Why Selenium Over BeautifulSoup?

**Challenge:** Noon.com is a modern e-commerce platform with heavy JavaScript rendering and dynamic content loading.

**Decision:** Selenium WebDriver

**Rationale:**
- **JavaScript Execution:** Product prices, seller information, and modals are loaded dynamically via JavaScript. BeautifulSoup can only parse static HTML.
- **Multi-Seller Modal:** The "Other Sellers" button triggers a JavaScript modal that doesn't exist in the initial HTML. Selenium can click buttons and wait for dynamic content.
- **Anti-Scraping Measures:** Noon.com detects headless browsers. Selenium allows us to:
  - Set realistic user agents
  - Disable automation flags
  - Add random delays to mimic human behavior
  - Scroll and interact naturally

**Trade-offs Considered:**
- **Speed:** Selenium is slower than BeautifulSoup, but accuracy and completeness are prioritized
- **Resource Usage:** Higher memory/CPU usage, but necessary for JavaScript-heavy sites
- **Maintenance:** CSS selectors may change, but the modular `config.py` makes updates easy

#### Assignment 1: Why Modular Architecture?

**Structure:** Separated into `main.py`, `noon_scraper.py`, `excel_exporter.py`, and `config.py`

**Benefits:**
- **Maintainability:** Each module has a single responsibility
- **Testability:** Individual components can be tested in isolation
- **Reusability:** `NoonScraper` class can be imported and used in other projects
- **Configuration:** CSS selectors and timeouts centralized in `config.py` for easy updates

#### Assignment 2: Why Pandas Over Manual Processing?

**Challenge:** Merge two different data sources with complex transformations.

**Decision:** Pandas DataFrames

**Rationale:**
- **Data Transformation:** Built-in functions for date parsing, calculations, and type conversions
- **Performance:** Vectorized operations are faster than row-by-row loops for large datasets
- **Excel Integration:** Native support for reading/writing Excel with `openpyxl` engine
- **Data Validation:** Easy to handle missing values, type mismatches, and edge cases

#### Assignment 2: Why Dynamic Mapping?

**Challenge:** Create a reusable script with zero hardcoding.

**Decision:** Configuration-driven approach using "Column Relations Sheet"

**Benefits:**
- **Reusability:** Same script works for any Excel file with the same structure
- **Flexibility:** Business users can update mappings without touching code
- **Maintainability:** Adding new transformations doesn't require code changes
- **Scalability:** Can handle new columns or data sources by updating the config sheet

### Design Patterns Used

1. **Object-Oriented Design (OOP)**
   - `NoonScraper` and `ReportMerger` classes encapsulate related functionality
   - Private methods (prefixed with `_`) hide implementation details
   - Public methods provide clean interfaces

2. **Error Handling**
   - Try-except blocks prevent single failures from crashing entire process
   - Graceful degradation (e.g., "N/A" for missing data)
   - Comprehensive logging for debugging

3. **Configuration Pattern**
   - Externalized configuration in `config.py` and Excel sheets
   - Easy to modify behavior without code changes
   - Supports different environments (headless vs. GUI mode)

4. **Progress Feedback**
   - `tqdm` progress bars show real-time status
   - Logging provides detailed execution history
   - User-friendly console output

### Performance Optimizations

#### Assignment 1: Scraping Efficiency
- **Smart Waiting:** Uses explicit waits instead of fixed `time.sleep()` where possible
- **Batch Processing:** Collects all product URLs before scraping to optimize navigation
- **Minimal Requests:** Extracts maximum data per page visit
- **Resource Cleanup:** Properly closes browser to free memory

#### Assignment 2: Data Processing Speed
- **Vectorized Operations:** Uses pandas built-in functions instead of Python loops
- **Memory Efficiency:** Processes data in chunks where applicable
- **Minimal I/O:** Reads Excel file once, writes once
- **Progress Tracking:** `tqdm` adds negligible overhead while improving UX

### Scalability Considerations

**Assignment 1:**
- Can handle 100+ products per keyword
- Supports multiple keywords in single run
- Configurable limits to control execution time
- Logging helps identify bottlenecks

**Assignment 2:**
- Handles 1000+ rows efficiently
- Works with any number of columns
- Supports multiple transformation types
- Can be extended to handle additional data sources

---

## Quick Start

### Prerequisites

- Python 3.7 or higher
- Chrome browser (for Assignment 1)
- Internet connection

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/Tatendamarimo/revent-automation-assignment.git
   cd revent-automation-assignment
   ```

2. **Install dependencies for Assignment 1:**
   ```bash
   cd revent-assignment-1-noon-scraper
   pip install -r requirements.txt
   ```

3. **Install dependencies for Assignment 2:**
   ```bash
   cd ../revent-assignment-2-report-merger
   pip install -r requirements.txt
   ```

### Running the Scripts

**Assignment 1 - Noon Scraper:**
```bash
cd revent-assignment-1-noon-scraper
python main.py
```

**Assignment 2 - Report Merger:**
```bash
cd revent-assignment-2-report-merger
python report_merger.py
```

---

## Output Files

### Assignment 1 Output
- **Location:** `revent-assignment-1-noon-scraper/output/`
- **Format:** Excel (.xlsx)
- **Naming:** `noon_scraper_<keyword>_<timestamp>.xlsx`
- **Sample:** `noon_scraper_iphone_2026-01-09_16-23-59.xlsx`

**Data Fields:**
1. Search Keyword
2. Category
3. Title
4. Description
5. Price
6. Rating
7. Reviews
8. Seller
9. Product URL

### Assignment 2 Output
- **Location:** Same directory as input file
- **Format:** Excel (.xlsx)
- **Naming:** `<original_filename>_MERGED.xlsx`
- **Sample:** `Assignments- Sheet Automation_MERGED.xlsx`

**Sheets Included:**
- Summary Sheet (merged data)
- Amazon (original)
- Noon (original)
- Column Relations Sheet (mapping)

---

## Configuration

### Assignment 1
Edit `config.py` to customize:
- Timeouts and delays
- CSS selectors
- Output directory
- Filename format

### Assignment 2
Use the "Column Relations Sheet" in your Excel file to configure:
- Column mappings
- Transformation rules
- Data processing logic

---

## Logic Explanation

### Assignment 1: Noon Scraper Logic

1. **Initialization**
   - Sets up Selenium WebDriver with Chrome
   - Configures anti-detection measures (user agent, automation flags)
   - Initializes logging system

2. **Search Process**
   - Navigates to noon.com
   - Enters search keyword
   - Waits for product grid to load

3. **Product Extraction**
   - Scrolls through product listings
   - Extracts basic product information
   - Clicks "Other Sellers" button for multi-seller products

4. **Seller Detection**
   - Opens seller modal
   - Extracts all available sellers
   - Captures seller-specific pricing

5. **Data Processing**
   - Combines product and seller data
   - Creates one row per product-seller combination
   - Handles missing data gracefully

6. **Export**
   - Formats data using pandas
   - Exports to Excel with proper column widths
   - Saves with timestamp in filename

### Assignment 2: Report Merger Logic

1. **Load Data**
   - Reads Amazon, Noon, and Column Relations sheets
   - Validates data structure

2. **Parse Mapping**
   - Creates dynamic column mapping dictionaries
   - Extracts transformation rules from Remarks column

3. **Transform Data**
   - **Date Extraction:** Splits dates into day, month, year
   - **Calculations:** Multiplies price × quantity for totals
   - **Conditional Logic:** Applies channel/contract-specific rules
   - **NA Marking:** Sets specified fields to "NA"

4. **Merge & Export**
   - Combines Amazon and Noon data
   - Adds source tracking column
   - Generates Summary Sheet
   - Saves complete workbook

---

## Libraries Used

### Assignment 1: Noon Scraper
| Library | Version | Purpose |
|---------|---------|---------|
| selenium | Latest | Web browser automation and JavaScript handling |
| webdriver-manager | Latest | Automatic ChromeDriver installation and updates |
| pandas | Latest | Data manipulation and DataFrame operations |
| openpyxl | Latest | Excel file creation and formatting |
| tqdm | Latest | Progress bars for real-time scraping feedback |

### Assignment 2: Report Merger
| Library | Version | Purpose |
|---------|---------|---------|
| pandas | ≥1.3.0 | Data manipulation, merging, and transformations |
| openpyxl | ≥3.0.0 | Excel file reading and writing |
| tqdm | Latest | Progress bars for data processing feedback |

---

## Troubleshooting

### Assignment 1 Issues

**ChromeDriver not found:**
- Ensure Chrome browser is installed
- The script auto-downloads ChromeDriver

**CAPTCHA appears:**
- Increase delays in `config.py`
- The script includes random delays to avoid detection

**No data scraped:**
- Check internet connection
- Verify noon.com is accessible
- Review `scraper.log` for errors

### Assignment 2 Issues

**File not found:**
- Ensure Excel file is in the same directory
- Or provide full file path as argument

**Column mapping errors:**
- Verify "Column Relations Sheet" exists
- Check column names match exactly

**Transformation not working:**
- Review Remarks column syntax
- Check `report_merger.py` for supported transformations

