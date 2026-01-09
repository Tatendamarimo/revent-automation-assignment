"""
Noon Web Scraper - Main Entry Point
Author: Revent Automation Assignment
Description: Scrapes product listings from noon.com with multi-seller support
"""

import sys
from noon_scraper import NoonScraper
from excel_exporter import ExcelExporter
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def print_banner():
    """Print application banner"""
    banner = """
    ╔══════════════════════════════════════════════════════════╗
    ║                                                          ║
    ║              NOON.COM WEB SCRAPER v1.0                   ║
    ║                                                          ║
    ║  Scrapes product listings with multi-seller support     ║
    ║  Exports to Excel format                                ║
    ║                                                          ║
    ╚══════════════════════════════════════════════════════════╝
    """
    print(banner)


def get_user_input():
    """
    Get search keywords from user
    
    Returns:
        list: List of keywords
    """
    print("\n" + "="*60)
    print("KEYWORD INPUT")
    print("="*60)
    print("\nEnter search keywords (one per line)")
    print("Press Enter twice when done, or type 'done' to finish\n")
    
    keywords = []
    while True:
        keyword = input(f"Keyword {len(keywords) + 1}: ").strip()
        
        if keyword.lower() == 'done' or keyword == '':
            if keywords:
                break
            else:
                print("Please enter at least one keyword!")
                continue
        
        keywords.append(keyword)
    
    return keywords


def get_max_products():
    """
    Ask user for maximum products per keyword
    
    Returns:
        int or None: Max products or None for all
    """
    print("\n" + "="*60)
    print("PRODUCT LIMIT")
    print("="*60)
    
    while True:
        response = input("\nLimit products per keyword? (Enter number or press Enter for all): ").strip()
        
        if response == '':
            return None
        
        try:
            max_products = int(response)
            if max_products > 0:
                return max_products
            else:
                print("Please enter a positive number!")
        except ValueError:
            print("Invalid input! Please enter a number or press Enter.")


def confirm_scrape(keywords, max_products):
    """
    Show summary and confirm before scraping
    
    Args:
        keywords (list): List of keywords
        max_products (int or None): Max products limit
        
    Returns:
        bool: True if confirmed
    """
    print("\n" + "="*60)
    print("SCRAPING SUMMARY")
    print("="*60)
    print(f"\nKeywords to scrape: {len(keywords)}")
    for idx, kw in enumerate(keywords, 1):
        print(f"  {idx}. {kw}")
    
    limit_text = f"{max_products} products" if max_products else "All products"
    print(f"\nProducts per keyword: {limit_text}")
    print(f"\nEstimated time: ~{len(keywords) * (max_products or 20) * 5} seconds")
    
    print("\n" + "="*60)
    
    while True:
        response = input("\nProceed with scraping? (yes/no): ").strip().lower()
        if response in ['yes', 'y']:
            return True
        elif response in ['no', 'n']:
            return False
        else:
            print("Please enter 'yes' or 'no'")


def main():
    """Main application flow"""
    
    # Print banner
    print_banner()
    
    try:
        # Get user input
        keywords = get_user_input()
        max_products = get_max_products()
        
        # Confirm before proceeding
        if not confirm_scrape(keywords, max_products):
            print("\nScraping cancelled by user")
            return
        
        # Initialize scraper
        print("\n" + "="*60)
        print("INITIALIZING SCRAPER")
        print("="*60 + "\n")
        
        scraper = NoonScraper(headless=False)  # Set to True for headless mode
        
        # Start scraping
        print("\n" + "="*60)
        print("SCRAPING IN PROGRESS")
        print("="*60 + "\n")
        print("Please wait... This may take several minutes.\n")
        
        data = scraper.scrape(keywords, max_products_per_keyword=max_products)
        
        # Check if data was scraped
        if not data:
            print("\nNo data was scraped. Please check the logs for errors.")
            return
        
        # Export to Excel
        print("\n" + "="*60)
        print("EXPORTING TO EXCEL")
        print("="*60 + "\n")
        
        exporter = ExcelExporter()
        
        # Use first keyword for filename if only one keyword
        filename_keyword = keywords[0] if len(keywords) == 1 else None
        filepath = exporter.export_to_excel(data, keyword=filename_keyword)
        
        # Success message
        print("\n" + "="*60)
        print("SCRAPING COMPLETED SUCCESSFULLY!")
        print("="*60)
        print(f"\nTotal products scraped: {len(data)}")
        print(f"Excel file saved: {filepath}")
        print(f"\nTip: Open the Excel file to view all scraped data")
        print("="*60 + "\n")
        
    except KeyboardInterrupt:
        print("\n\nScraping interrupted by user (Ctrl+C)")
        sys.exit(1)
    
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}", exc_info=True)
        print(f"\nAn error occurred: {str(e)}")
        print("Check scraper.log for details")
        sys.exit(1)


if __name__ == "__main__":
    main()
