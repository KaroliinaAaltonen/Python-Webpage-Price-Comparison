import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# Function for getting URL HTML content
def get_html_from_url(url):
    chrome_options = Options()
    chrome_options.add_argument('--headless')  # Run Chrome in headless mode (without UI)

    service = ChromeService(executable_path=ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        driver.get(url)
        time.sleep(10)  # Allow time for JavaScript to load (you can adjust as needed)
        return driver.page_source
    finally:
        driver.quit()

# Searching product's price IN PRISMA
# Prisma's search engine Supports product code, name
def search_prisma(parameter):
    url = f"https://www.prisma.fi/haku?search={parameter}"
    html = get_html_from_url(url)

    if html:
        soup = BeautifulSoup(html, 'html.parser')
        
        print("\nPRISMA:")
        no_results = prisma_search_result_check(soup)
        if no_results:
            print("Ei tuloksia haulla:",parameter)
            return 0
        price = prisma_extract_price(soup)
        product_name = prisma_extract_product_name(soup)

        if product_code:
            print("tuotekoodi:",parameter)
        if product_name:
            print("tuotenimi:",product_name)
        if price:
            print("hinta:",price,"€")

    else:
        print("Prisma: Error with handling HTML.")
    return 0

# Checking if product exists FOR PRISMA PRODUCTS
def prisma_search_result_check(soup):
    no_search_results_element = soup.find('p', {'data-test-id': 'no-search-results', 'class': 'search_noResultTitle__PTo4Q'})
    if no_search_results_element:
        result_text = no_search_results_element.get_text(strip=True)
        return result_text
    else:
        return None
    
# Extracting product name FOR PRISMA PRODUCTS
def prisma_extract_product_name(soup):
    # Find the <a> element with the specified class and data-test-id
    product_link = soup.find('a', class_='ProductCard_blockLink__NTUGB', attrs={'data-test-id': 'product-card-link'})

    # Extract the text content from the <a> element
    if product_link:
        product_name = product_link.get_text(strip=True)
        return product_name
    else:
        return None

# Extracting product price FOR PRISMA PRODUCTS
def prisma_extract_price(soup):
    # Find the <p> element with the specified class and data-test-id
    price_p = soup.find('p', class_='ProductCard_priceContainer__jUqDu', attrs={'data-test-id': 'product-card-price'})

    # Extract the text content from the <span> element
    if price_p:
        price_span = price_p.find('span', class_='ProductCard_finalPrice__xD9gs')
    if price_span:
        price_value = price_span.get_text(strip=True)
        return price_value
    else:
        return None


# Searching product's price IN PUUILO
# Puuilo's website supports EAN, product code, name
def search_puuilo(product_code):
    url = f"https://www.puuilo.fi/catalogsearch/result/?q={product_code}"
    html = get_html_from_url(url)

    if html:
        soup = BeautifulSoup(html, 'html.parser')
        price = puuilo_extract_price(soup)
        product_name = puuilo_extract_product_name(soup)
        
        print("\nPUUILO:")
        if product_code:
            print("tuotekoodi:",product_code)
        if product_name:
            print("tuotenimi:",product_name)
        if price:
            print("hinta:",price,"€")
    else:
        print("Puuilo: Error with handling HTML.")
    
    return 0

# Extracting product name FOR PUUILO PRODUCTS
# Dynamic, requires Selenium and ChromeDriver DOES NOT WORK YET
def puuilo_extract_product_name(soup):
    # Find the <h2> element with the specified class
    product_name_element = soup.find('h2', class_='product name product-item-name')
    if product_name_element:

    # Find the <a> element with the specified class
        product_name = product_name_element.find('a', class_='product-item-link').get_text(strip=True)
        return product_name
    else:
        return None
        
# Extracting product prices FOR PUUILO PRODUCTS
def puuilo_extract_price(soup):
    price_span = soup.find('span', class_='price')

    if price_span:
        price_text = price_span.get_text(strip=True).replace('\xa0', '')
        return price_text
    else:
        return None

# Searching product's price in K-RAUTA
# supports EAN/product code, name
def search_k_rauta(product_code):
    url = f"https://www.k-rauta.fi/etsi?query={product_code}"
    html = get_html_from_url(url)

    if html:
        soup = BeautifulSoup(html, 'html.parser')
        price = k_rauta_extract_price(soup)
        product_name = k_rauta_extract_product_name(soup)
        
        print("\nK-RAUTA:")
        if product_code:
            print("tuotekoodi:",product_code)
        if product_name:
            print("tuotenimi:",product_name)
        if price:
            print("hinta:",price)
    else:
        print("K-rauta: Error with handling HTML.")
    
    return 0

# Extracting product prices FOR K-RAUTA PRODUCTS
# Dynamic, requires Selenium and ChromeDriver
def k_rauta_extract_price(soup):
    # Find the <span> element with the specified class
    price_span = soup.find('span', class_='price-view-fi__sale-price--prefix')
    # Extract the text content from the <span> element
    if price_span:
        price_text = price_span.next_sibling.strip()
        return price_text
            
    else:
        print("Price <span> element not found.")
        
# Extracting product name FOR K-RAUTA PRODUCTS
# Dynamic, requires Selenium and ChromeDdriver
def k_rauta_extract_product_name(soup):
    # Find the <h2> element with the specified class and data-cy attribute
    product_name_h2 = soup.find('h2', class_='product-card__name', attrs={'data-cy': 'product-card-name'})

    # Extract the text content from the <h2> element
    if product_name_h2:
        product_name = product_name_h2.get_text(strip=True)
        return product_name
    else:
        return None


# Example usage:
ean_code = 6430050623548
product_code = 6435200187130

search_k_rauta(product_code)
search_prisma(product_code)
search_puuilo(product_code)

search_k_rauta(ean_code)
search_prisma(ean_code)
search_puuilo(ean_code)
