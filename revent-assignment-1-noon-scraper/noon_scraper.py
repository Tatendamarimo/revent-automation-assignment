"""
Noon Web Scraper Module
Handles scraping of product data from noon.com
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import random
import logging
from config import *

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format=LOG_FORMAT,
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class NoonScraper:
    """Main scraper class for noon.com"""
    
    def __init__(self, headless=False):
        """
        Initialize the scraper
        
        Args:
            headless (bool): Run browser in headless mode
        """
        self.headless = headless
        self.driver = None
        self.scraped_data = []
        
    def _init_driver(self):
        """Initialize Selenium WebDriver"""
        logger.info("Initializing Chrome WebDriver...")
        
        chrome_options = Options()
        if self.headless:
            chrome_options.add_argument('--headless')
        
        # Additional options to avoid detection
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--window-size=1920,1080')
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # User agent
        chrome_options.add_argument('user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
        
        logger.info("WebDriver initialized successfully")
    
    def _random_delay(self, min_delay=REQUEST_DELAY_MIN, max_delay=REQUEST_DELAY_MAX):
        """Add random delay to mimic human behavior"""
        delay = random.uniform(min_delay, max_delay)
        time.sleep(delay)
    
    def _safe_find_element(self, by, selector, timeout=ELEMENT_WAIT_TIMEOUT, parent=None):
        """
        Safely find element with wait
        
        Args:
            by: Selenium By locator
            selector: CSS selector or XPath
            timeout: Wait timeout
            parent: Parent element to search within
            
        Returns:
            WebElement or None
        """
        try:
            element_source = parent if parent else self.driver
            element = WebDriverWait(element_source, timeout).until(
                EC.presence_of_element_located((by, selector))
            )
            return element
        except TimeoutException:
            return None
    
    def _safe_find_elements(self, by, selector, timeout=ELEMENT_WAIT_TIMEOUT, parent=None):
        """
        Safely find multiple elements
        
        Args:
            by: Selenium By locator
            selector: CSS selector or XPath
            timeout: Wait timeout
            parent: Parent element to search within
            
        Returns:
            List of WebElements
        """
        try:
            element_source = parent if parent else self.driver
            WebDriverWait(element_source, timeout).until(
                EC.presence_of_element_located((by, selector))
            )
            return element_source.find_elements(by, selector)
        except TimeoutException:
            return []
    
    def _get_text_safe(self, element, selector=None):
        """
        Safely extract text from element
        
        Args:
            element: WebElement or driver
            selector: Optional CSS selector
            
        Returns:
            str: Extracted text or empty string
        """
        try:
            if selector:
                target = element.find_element(By.CSS_SELECTOR, selector)
            else:
                target = element
            return target.text.strip()
        except:
            return ""
    
    def search_keyword(self, keyword):
        """
        Search for a keyword on noon.com
        
        Args:
            keyword (str): Search keyword
            
        Returns:
            bool: Success status
        """
        try:
            logger.info(f"Searching for keyword: {keyword}")
            
            # Navigate to search URL
            search_url = f"{SEARCH_URL}{keyword}"
            self.driver.get(search_url)
            
            # Wait for results to load
            self._random_delay(2, 3)
            
            # Check if results loaded
            results = self._safe_find_elements(By.CSS_SELECTOR, SELECTORS['product_card'], timeout=10)
            
            if results:
                logger.info(f"Found {len(results)} products on first page")
                return True
            else:
                logger.warning(f"No results found for keyword: {keyword}")
                return False
                
        except Exception as e:
            logger.error(f"Error searching for keyword '{keyword}': {str(e)}")
            return False
    
    def _extract_category(self):
        """Extract category from breadcrumbs"""
        try:
            breadcrumbs = self._safe_find_elements(By.CSS_SELECTOR, SELECTORS['breadcrumbs'])
            if breadcrumbs:
                # Get all breadcrumb texts except "Home"
                categories = [bc.text.strip() for bc in breadcrumbs if bc.text.strip() and bc.text.strip().lower() != 'home']
                return ' > '.join(categories) if categories else "N/A"
            return "N/A"
        except:
            return "N/A"
    
    def _extract_description(self):
        """Extract product description/highlights"""
        try:
            # Try to find highlights section
            highlights = self._safe_find_elements(By.CSS_SELECTOR, SELECTORS['highlights'])
            if highlights:
                desc_parts = [h.text.strip() for h in highlights[:3]]  # Get first 3 highlights
                return ' | '.join(desc_parts) if desc_parts else "N/A"
            return "N/A"
        except:
            return "N/A"
    
    def _extract_sellers(self, product_url):
        """
        Extract all sellers for a product
        
        Args:
            product_url (str): Product URL
            
        Returns:
            list: List of seller dictionaries
        """
        sellers = []
        
        try:
            # Wait for price element to load
            price_element = self._safe_find_element(By.CSS_SELECTOR, SELECTORS['price_now'], timeout=5)
            
            # Find primary seller
            primary_seller = self._get_text_safe(self.driver, SELECTORS['seller_name'])
            
            # Extract price using JavaScript textContent (more reliable than .text)
            if price_element:
                try:
                    price_text = self.driver.execute_script("return arguments[0].textContent;", price_element)
                    # Clean price text - remove any non-printable characters and extra whitespace
                    if price_text:
                        import re
                        # Remove Unicode special characters and keep only digits, decimal points, and basic ASCII
                        price_text = re.sub(r'[^\x20-\x7E]', '', price_text)  # Keep only printable ASCII
                        price_text = price_text.strip()
                        primary_price = f"AED {price_text}" if price_text else "N/A"
                    else:
                        primary_price = "N/A"
                except:
                    primary_price = "N/A"
            else:
                primary_price = "N/A"
            
            primary_rating = self._get_text_safe(self.driver, SELECTORS['rating_value'])
            
            # Log price extraction for debugging
            if primary_price == "N/A":
                logger.warning(f"Price not found for product. Tried selector: {SELECTORS['price_now']}")
            else:
                logger.info(f"Extracted price: {primary_price}")
            
            # Always add primary seller (default to 'noon' if not found)
            sellers.append({
                'name': primary_seller if primary_seller else 'noon',
                'price': primary_price,
                'rating': primary_rating
            })
            
            # Check for other sellers button
            try:
                # Look for "other seller" button using XPath
                other_sellers_buttons = self.driver.find_elements(By.XPATH, "//button[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'other seller')]")
                
                if other_sellers_buttons:
                    logger.info(f"Found 'Other Sellers' button, clicking to view all sellers...")
                    
                    # Click the button
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", other_sellers_buttons[0])
                    time.sleep(0.5)
                    other_sellers_buttons[0].click()
                    
                    # Wait for modal to appear
                    time.sleep(2)
                    
                    # Extract sellers from modal
                    modal_sellers = self._safe_find_elements(By.CSS_SELECTOR, SELECTORS['modal_sellers'], timeout=5)
                    
                    logger.info(f"Found {len(modal_sellers)} seller cards in modal")
                    
                    for seller_card in modal_sellers:
                        try:
                            seller_name = self._get_text_safe(seller_card, SELECTORS['modal_seller_name'])
                            seller_price = self._get_text_safe(seller_card, SELECTORS['modal_seller_price'])
                            seller_rating = self._get_text_safe(seller_card, SELECTORS['modal_seller_rating'])
                            
                            # Avoid duplicates
                            if seller_name and not any(s['name'] == seller_name for s in sellers):
                                sellers.append({
                                    'name': seller_name,
                                    'price': seller_price if seller_price else primary_price,
                                    'rating': seller_rating if seller_rating else 'N/A'
                                })
                        except:
                            continue
                    
                    # Close modal
                    try:
                        close_buttons = self.driver.find_elements(By.CSS_SELECTOR, SELECTORS['close_modal'])
                        if close_buttons:
                            close_buttons[0].click()
                            time.sleep(0.5)
                        else:
                            # Press ESC key
                            self.driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                            time.sleep(0.5)
                    except:
                        pass
                        
            except Exception as e:
                logger.debug(f"No other sellers found or error accessing modal: {str(e)}")
        
        except Exception as e:
            logger.error(f"Error extracting sellers: {str(e)}")
        
        # If no sellers found, add default
        if not sellers:
            sellers.append({
                'name': 'noon',
                'price': 'N/A',
                'rating': 'N/A'
            })
        
        return sellers
    
    def scrape_product_details(self, product_url, keyword):
        """
        Scrape details from a single product page
        
        Args:
            product_url (str): URL of product
            keyword (str): Search keyword used
            
        Returns:
            list: List of product data dictionaries (one per seller)
        """
        product_data_list = []
        
        try:
            logger.info(f"Scraping product: {product_url}")
            
            # Navigate to product page
            self.driver.get(product_url)
            self._random_delay(PRODUCT_DETAIL_DELAY, PRODUCT_DETAIL_DELAY + 1)
            
            # Scroll to load content and wait for price to appear
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight/2);")
            time.sleep(2)  # Increased wait time for dynamic content
            
            # Extract common product information
            title = self._get_text_safe(self.driver, SELECTORS['product_title'])
            category = self._extract_category()
            description = self._extract_description()
            
            # Extract rating and reviews
            rating_element = self._safe_find_element(By.CSS_SELECTOR, SELECTORS['rating_value'])
            rating = self._get_text_safe(rating_element) if rating_element else "N/A"
            
            reviews_element = self._safe_find_element(By.CSS_SELECTOR, SELECTORS['reviews_count'])
            reviews = self._get_text_safe(reviews_element) if reviews_element else "N/A"
            
            # Extract all sellers (includes price extraction)
            sellers = self._extract_sellers(product_url)
            
            # Create a row for each seller
            for seller in sellers:
                product_data = {
                    'Search Keyword': keyword,
                    'Category': category,
                    'Title': title,
                    'Description': description,
                    'Price': seller['price'],
                    'Rating': rating,
                    'Reviews': reviews,
                    'Seller': seller['name'],
                    'Product URL': product_url
                }
                product_data_list.append(product_data)
            
            logger.info(f"Scraped product with {len(sellers)} seller(s)")
            
        except Exception as e:
            logger.error(f"Error scraping product {product_url}: {str(e)}")
        
        return product_data_list
    
    def scrape_search_results(self, keyword, max_products=None):
        """
        Scrape all products from search results
        
        Args:
            keyword (str): Search keyword
            max_products (int): Maximum number of products to scrape (None for all)
            
        Returns:
            list: List of all scraped product data
        """
        all_data = []
        
        try:
            # Search for keyword
            if not self.search_keyword(keyword):
                return all_data
            
            # Scroll to load more products
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight/3);")
            time.sleep(1)
            
            # Try multiple selectors to find product links
            product_urls = []
            
            # Method 1: Try finding product cards first
            product_cards = self._safe_find_elements(By.CSS_SELECTOR, SELECTORS['product_card'], timeout=5)
            logger.info(f"Found {len(product_cards)} product cards")
            
            if product_cards:
                for card in product_cards:
                    try:
                        # Try to find link within card
                        link = card if card.tag_name == 'a' else card.find_element(By.TAG_NAME, 'a')
                        url = link.get_attribute('href')
                        if url and '/p/' in url and url not in product_urls:
                            product_urls.append(url)
                    except:
                        continue
            
            # Method 2: If no URLs found, try finding all links with /p/ in href
            if not product_urls:
                logger.info("Trying alternative method to find product links...")
                all_links = self.driver.find_elements(By.TAG_NAME, 'a')
                for link in all_links:
                    try:
                        url = link.get_attribute('href')
                        if url and '/p/' in url and '/uae-en/' in url and url not in product_urls:
                            product_urls.append(url)
                    except:
                        continue
            
            if not product_urls:
                logger.warning("No product URLs found on search results page")
                return all_data
            
            # Limit if max_products specified
            if max_products:
                product_urls = product_urls[:max_products]
            
            logger.info(f"Found {len(product_urls)} products to scrape")
            
            # Scrape each product
            for idx, url in enumerate(product_urls, 1):
                logger.info(f"Progress: {idx}/{len(product_urls)}")
                
                product_data = self.scrape_product_details(url, keyword)
                all_data.extend(product_data)
                
                # Random delay between products
                self._random_delay()
            
            logger.info(f"Completed scraping {len(product_urls)} products")
            
        except Exception as e:
            logger.error(f"Error scraping search results: {str(e)}")
        
        return all_data
    
    def scrape(self, keywords, max_products_per_keyword=None):
        """
        Main scraping method
        
        Args:
            keywords (list or str): Keyword(s) to search
            max_products_per_keyword (int): Max products per keyword
            
        Returns:
            list: All scraped data
        """
        # Convert single keyword to list
        if isinstance(keywords, str):
            keywords = [keywords]
        
        try:
            # Initialize driver
            self._init_driver()
            
            # Scrape each keyword
            for keyword in keywords:
                logger.info(f"\n{'='*60}")
                logger.info(f"Starting scrape for keyword: '{keyword}'")
                logger.info(f"{'='*60}\n")
                
                data = self.scrape_search_results(keyword, max_products_per_keyword)
                self.scraped_data.extend(data)
                
                logger.info(f"Scraped {len(data)} rows for keyword '{keyword}'")
            
            logger.info(f"\n{'='*60}")
            logger.info(f"SCRAPING COMPLETED")
            logger.info(f"Total rows scraped: {len(self.scraped_data)}")
            logger.info(f"{'='*60}\n")
            
        except Exception as e:
            logger.error(f"Error during scraping: {str(e)}")
        
        finally:
            # Close driver
            if self.driver:
                self.driver.quit()
                logger.info("Browser closed")
        
        return self.scraped_data
    
    def get_data(self):
        """Return scraped data"""
        return self.scraped_data
