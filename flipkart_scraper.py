import requests
from bs4 import BeautifulSoup
import time
import random
import logging
from urllib.parse import urljoin
from requests.exceptions import RequestException

logger = logging.getLogger(__name__)

def search(keywords, num_products=30):
    try:
        all_data = fetch_flipkart_data(keywords, num_products)
        products = process_flipkart_data(all_data, num_products)
        if not products:
            error_msg = f"No products found for '{keywords}'"
            logger.error(error_msg)
            raise Exception(error_msg)
        return products
    except Exception as e:
        error_msg = f"Error during Flipkart search: {str(e)}"
        logger.error(error_msg)
        raise Exception(error_msg)

def fetch_flipkart_data(keyword, num_products):
    all_data = []
    url = f"https://www.flipkart.com/search?q={keyword.replace(' ', '+')}&otracker=search&otracker1=search&marketplace=FLIPKART&as-show=off&as=off"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3",
        "Accept-Language": "en-US,en;q=0.5",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
    }

    page = 1
    while len(all_data) * 24 < num_products:  # Assuming 24 products per page
        try:
            logger.info(f"Fetching page {page} for '{keyword}'")
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, "lxml")
            all_data.append(soup)
            
            next_page = soup.find("a", class_="_9QVEpD")
            if next_page and "href" in next_page.attrs:
                url = urljoin("https://www.flipkart.com", next_page["href"])
                page += 1
                time.sleep(random.uniform(4,6))
            else:
                logger.info(f"No more pages found for '{keyword}'")
                break
        except RequestException as e:
            logger.error(f"An error occurred while fetching results for '{keyword}' on page {page}: {e}")
            break

    return all_data

def process_flipkart_data(all_data, num_products=30):
    products = []
    for soup in all_data:
        product_containers = soup.find_all("div", attrs={"data-id": True})
        
        for container in product_containers:
            if len(products) >= num_products:
                break

            try:
                product_id = container['data-id']
                
                # Check for name in multiple possible elements
                name_elem = container.find("a", class_="wjcEIp") or container.find("div", class_="KzDlHZ")
                name = name_elem.get("title", name_elem.text.strip()) if name_elem else "N/A"
                
                price_elem = container.find("div", class_="Nx9bqj")
                
                # Check for link in multiple possible elements
                link_elem = container.find("a", class_="wjcEIp") or container.find("a", class_="CGtC98")
                rating_elem = container.find("div", class_="XQDdHH")
                reviews_elem = container.find("span", class_="Wphh3N")
                
                price = price_elem.get_text(strip=True) if price_elem else "N/A"
                link = f"https://www.flipkart.com{link_elem['href']}" if link_elem and 'href' in link_elem.attrs else "N/A"
                rating = rating_elem.get_text(strip=True) if rating_elem else "N/A"
                
                # Extract only the number of ratings
                reviews = reviews_elem.get_text(strip=True).strip("()") if reviews_elem else "N/A"
                if reviews != "N/A":
                    reviews = reviews.split()[0]  # Take only the first part (number of ratings)
                
                product = {
                    "rank": len(products) + 1,
                    "product_id": product_id,
                    "title": name,
                    "price": price,
                    "link": link,
                    "rating": rating,
                    "reviews": reviews
                }
                
                logger.info(f"Processed product: {product}")
                products.append(product)
            except Exception as e:
                logger.error(f"Error processing product: {str(e)}")
                logger.error(f"Product HTML: {container}")

    logger.info(f"Processed {len(products)} products")
    return products[:num_products]