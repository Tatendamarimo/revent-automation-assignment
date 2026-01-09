# Quick Start Guide

## For Reviewers/Evaluators

### Running Assignment 1: Noon Web Scraper

```bash
cd revent-assignment-1-noon-scraper
pip install -r requirements.txt
python main.py
```

**Test with:**
- Keyword: `iphone`
- Limit: `5` products

**Expected output:** Excel file in `output/` folder with ~5-15 rows (depending on sellers)

**Time:** ~2-3 minutes

---

### Running Assignment 2: Report Merger

```bash
cd revent-assignment-2-report-merger
pip install -r requirements.txt
python report_merger.py
```

**Expected output:** `Assignments- Sheet Automation_MERGED.xlsx` with Summary Sheet

**Time:** ~5 seconds

---

## Troubleshooting

**Assignment 1:**
- If ChromeDriver fails: Make sure Chrome browser is installed
- If no data scraped: Check internet connection
- If CAPTCHA appears: Increase delays in `config.py`

**Assignment 2:**
- If file not found: Ensure you're in the correct directory
- If column errors: Check Excel file has required sheets

---

## What to Look For

### Assignment 1 Highlights
✓ Browser opens and navigates to noon.com  
✓ Progress bar shows real-time scraping status  
✓ Clicks "Other Sellers" button when available  
✓ Handles errors gracefully (continues on failure)  
✓ Excel output has all 9 required fields  
✓ Multiple sellers = multiple rows  

### Assignment 2 Highlights
✓ Progress bars for Noon and Amazon processing  
✓ Automatic date component extraction  
✓ Price × Quantity calculations  
✓ Source column shows data origin  
✓ All original sheets preserved  
✓ Works with any similar Excel file (zero hardcoding)  

---

## Questions?

Check the detailed README files in each assignment folder or review `scraper.log` for debugging information.
