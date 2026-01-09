"""
Configuration settings for Noon Web Scraper
"""

# Base URL
BASE_URL = "https://www.noon.com/uae-en"
SEARCH_URL = f"{BASE_URL}/search/?q="

# Timeout settings (in seconds)
PAGE_LOAD_TIMEOUT = 30
ELEMENT_WAIT_TIMEOUT = 10
SCROLL_PAUSE_TIME = 2

# Delays to avoid bot detection (in seconds)
REQUEST_DELAY_MIN = 2
REQUEST_DELAY_MAX = 4
PRODUCT_DETAIL_DELAY = 1.5

# CSS Selectors (using partial matching for dynamic class names)
SELECTORS = {
    # Search results page
    'product_links': 'a[class*="productContainer"]',
    'product_card': 'div[data-qa="product-card"], div[class*="productContainer"], a[class*="productBoxLink"]',
    
    # Product detail page
    'product_title': 'h1',
    'price_now': '[class*="priceNowText"], [class*="priceNow"]',
    'rating_value': 'div[class*="rating"] span, span[class*="rating"]',
    'reviews_count': '[class*="reviews"], [class*="rating"]',
    'breadcrumbs': 'nav[class*="breadcrumb"] a, [class*="Breadcrumb"] a',
    'seller_name': '[class*="soldBy"], [class*="SoldBy"]',
    'other_sellers_button': 'button:has-text("other seller")',
    'highlights': 'ul[class*="highlights"] li, div[class*="highlights"] li, [class*="description"]',
    
    # Other sellers modal
    'modal_sellers': '[class*="offerCard"], [class*="sellerCard"], div[class*="offer"]',
    'modal_seller_name': '[class*="sellerName"], [class*="partner"] strong, strong',
    'modal_seller_price': '[class*="price"]',
    'modal_seller_rating': '[class*="rating"]',
    'close_modal': 'button[class*="close"], [aria-label="Close"], button[class*="Close"]',
}

# Output settings
OUTPUT_DIR = "output"
OUTPUT_FILENAME_PREFIX = "noon_scraper"

# Excel column headers
EXCEL_HEADERS = [
    "Search Keyword",
    "Category",
    "Title",
    "Description",
    "Price",
    "Rating",
    "Reviews",
    "Seller",
    "Product URL"
]

# Logging
LOG_FILE = "scraper.log"
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
