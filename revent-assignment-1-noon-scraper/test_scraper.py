"""
Test script for Noon Web Scraper
Tests basic functionality with a limited scrape
"""

from noon_scraper import NoonScraper
from excel_exporter import ExcelExporter
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def test_scraper():
    """Test the scraper with a single keyword and limited products"""
    
    print("\n" + "="*60)
    print("NOON WEB SCRAPER - TEST MODE")
    print("="*60)
    print("\nTesting with keyword: 'iphone'")
    print("Limiting to 3 products for quick test\n")
    
    try:
        # Initialize scraper
        print("Initializing scraper...")
        scraper = NoonScraper(headless=False)
        
        # Scrape
        print("Starting scrape...\n")
        data = scraper.scrape(keywords=['iphone'], max_products_per_keyword=3)
        
        # Check results
        if data:
            print(f"\nSuccessfully scraped {len(data)} rows")
            
            # Display sample data
            print("\nSample data (first row):")
            print("-" * 60)
            for key, value in data[0].items():
                print(f"{key}: {value}")
            print("-" * 60)
            
            # Export to Excel
            print("\nExporting to Excel...")
            exporter = ExcelExporter()
            filepath = exporter.export_to_excel(data, keyword='iphone_test')
            
            print(f"\nTEST PASSED!")
            print(f"Excel file created: {filepath}")
            
        else:
            print("\nTEST FAILED: No data scraped")
            
    except Exception as e:
        print(f"\nTEST FAILED: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_scraper()
