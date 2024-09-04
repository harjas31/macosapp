import requests
from bs4 import BeautifulSoup
import logging
import re
from requests.exceptions import RequestException

logger = logging.getLogger(__name__)

def fetch_amazon_product_info(identifier):
    try:
        soup = fetch_amazon_data(identifier)
        return process_amazon_data(soup, identifier)
    except Exception as e:
        logger.error(f"Error fetching Amazon product info: {str(e)}")
        return None

def fetch_amazon_data(identifier):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
    }

    if re.match(r'^[A-Z0-9]{10}$', identifier):
        url = f"https://www.amazon.in/dp/{identifier}"
    elif 'amazon.in' in identifier:
        url = identifier
    else:
        raise ValueError(f"Invalid Amazon URL or ASIN: {identifier}")

    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        return BeautifulSoup(response.content, 'html.parser')
    except RequestException as e:
        logger.error(f"Error fetching Amazon product data: {str(e)}")
        raise

def process_amazon_data(soup, identifier):
    try:
        title_elem = soup.find("span", {"id": "productTitle"})
        title = title_elem.text.strip() if title_elem else "N/A"
        
        price_elem = soup.find("span", {"class": "a-price-whole"})
        price = price_elem.text.strip() if price_elem else "N/A"
        
        rating_elem = soup.find("span", {"class": "a-icon-alt"})
        rating = rating_elem.text.split()[0] if rating_elem else "N/A"
        
        reviews_elem = soup.find("span", {"id": "acrCustomerReviewText"})
        reviews = reviews_elem.text.split()[0] if reviews_elem else "N/A"

        # Check stock availability
        stock_status = check_stock_availability(soup)

        # Updated bestseller rank logic
        bestseller_ranks = []
        rank_elem = soup.find("div", {"id": "detailBulletsWrapper_feature_div"})
        if rank_elem:
            rank_items = rank_elem.find_all("span", {"class": "a-list-item"})
            for item in rank_items:
                if "Best Sellers Rank" in item.text:
                    rank_text = item.text.strip()
                    ranks = re.findall(r'#([\d,]+) in ([^(#]+)', rank_text)
                    for rank, category in ranks:
                        bestseller_ranks.append(f"#{rank.replace(',', '')} in {category.strip()}")

        # If the above method doesn't work, try an alternative approach
        if not bestseller_ranks:
            rank_table = soup.find("table", {"id": "productDetails_detailBullets_sections1"})
            if rank_table:
                rank_rows = rank_table.find_all("tr")
                for row in rank_rows:
                    if "Best Sellers Rank" in row.text:
                        rank_text = row.find("td", {"class": "a-size-base"}).text.strip()
                        ranks = re.findall(r'#([\d,]+) in ([^(#]+)', rank_text)
                        for rank, category in ranks:
                            bestseller_ranks.append(f"#{rank.replace(',', '')} in {category.strip()}")

        asin = identifier if not identifier.startswith('http') else re.search(r'/dp/([A-Z0-9]{10})', identifier).group(1)

        return {
            "ASIN": asin,
            "title": title,
            "price": price,
            "rating": rating,
            "reviews": reviews,
            "link": f"https://www.amazon.in/dp/{asin}",
            "BestSeller": " | ".join(bestseller_ranks) if bestseller_ranks else "N/A",
            "In Stock": stock_status
        }
    except Exception as e:
        logger.error(f"Error processing Amazon product data: {str(e)}")
        return None

def fetch_flipkart_product_info(url):
    try:
        soup = fetch_flipkart_data(url)
        return process_flipkart_data(soup, url)
    except Exception as e:
        logger.error(f"Error fetching Flipkart product info: {str(e)}")

def fetch_flipkart_data(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3",
        "Accept-Language": "en-US,en;q=0.5",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
    }

    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        return BeautifulSoup(response.content, 'lxml')
    except RequestException as e:
        logger.error(f"Error fetching Flipkart product data: {str(e)}")



def process_flipkart_data(soup, url):
    try:
        title_elem = soup.find('span', class_='VU-ZEz')
        title = title_elem.text.strip() if title_elem else 'N/A'
        price_elem = soup.find('div', class_='Nx9bqj CxhGGd')
        price = price_elem.text.strip() if price_elem else 'N/A'
        rating_elem = soup.find('div', class_='XQDdHH')
        rating = rating_elem.text.strip() if rating_elem else 'N/A'
        reviews_elem = soup.find('span', class_='Wphh3N')
        reviews = reviews_elem.text.strip() if reviews_elem else 'N/A'

        # Extract only the number of ratings
        if reviews != 'N/A':
            reviews = reviews.split()[0]

        return {
            "link": url,
            "title": title,
            "price": price,
            "rating": rating,
            "reviews": reviews,
        }
    except Exception as e:
        logger.error(f"Error processing Flipkart product data: {str(e)}")
def check_stock_availability(soup):
    try:
        # Check for "In stock" text
        stock_elem = soup.find("span", {"class": "a-size-medium a-color-success"})
        if stock_elem and "In stock" in stock_elem.text:
            return "Yes"
        if stock_elem and "Currently unavailable" in stock_elem.text:
            return "No"


        # If none of the above conditions are met, return "Unknown"
        return "Unknown"
    except Exception as e:
        logger.error(f"Error checking stock availability: {str(e)}")
        return "Unknown"