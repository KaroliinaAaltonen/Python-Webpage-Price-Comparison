import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
import concurrent.futures

from openpyxl.utils import get_column_letter

# web scraping functions
# driver is started locally (and ended)
# get_html_from_url() uses sleep=10 with puuilo.fi and tokmanni.fi as per their robots.txt
class Scraper:
    def __init__(self):
        self.driver = self.local_driver()

    def local_driver(self):
        chrome_options = Options()
        chrome_options.add_argument('--headless')  # Run Chrome in headless mode (without UI)
        service = ChromeService(executable_path=ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver

    def end_session(self):
        self.driver.quit()

    def get_html_from_url(self, url, sleep=1):
        try:
            self.driver.get(url)
            time.sleep(sleep)
            return self.driver.page_source
        except Exception as e:
            print(f"Error getting HTML from {url}: {str(e)}")
            return None

class PrismaScraper(Scraper):
    def __init__(self):
        super().__init__()
    def search_product(self, product_code):
        url = f"https://www.prisma.fi/haku?search={product_code}"
        html = self.get_html_from_url(url)

        if html:
            soup = BeautifulSoup(html, 'html.parser')
            
            print("\nPRISMA:")
            if self.no_search_result_check(soup):
                print("Ei tuloksia haulla:", product_code)
                return None
            
            price = self.extract_price(soup)
            product_name = self.extract_product_name(soup)
            brand_name = product_name.split(" ")[0].upper()

            if product_name and price:
                print("tuotekoodi:", product_code)
                print("tuotenimi:", product_name)
                print("hinta:", price, "€")
                return product_name, price, brand_name
        else:
            print("Prisma: Error with handling HTML.")
        return None

    def no_search_result_check(self, soup):
        no_search_results_element = soup.find('p', {'data-test-id': 'no-search-results', 'class': 'search_noResultTitle__PTo4Q'})
        if no_search_results_element:
            result_text = no_search_results_element.get_text(strip=True)
            return result_text
        else:
            return None

    def extract_product_name(self, soup):
        product_link = soup.find('a', class_='ProductCard_blockLink__NTUGB', attrs={'data-test-id': 'product-card-link'})
        if product_link:
            product_name = product_link.get_text(strip=True)
            return product_name
        else:
            return None

    def extract_price(self, soup):
        price_p = soup.find('p', class_='ProductCard_priceContainer__jUqDu', attrs={'data-test-id': 'product-card-price'})
        if price_p:
            price_span = price_p.find('span', class_='ProductCard_finalPrice__xD9gs')
            if price_span:
                price_value = price_span.get_text(strip=True)
                return price_value
        return None

class PuuiloScraper(Scraper):
    def __init__(self):
        super().__init__()
    def search_product(self, product_code):
        url = f"https://www.puuilo.fi/catalogsearch/result/?q={product_code}"
        html = self.get_html_from_url(url, 10)

        if html:
            soup = BeautifulSoup(html, 'html.parser')
            print("\nPUUILO:")
            
            no_results = self.search_result_check(soup)
            if no_results:
                print("Ei tuloksia haulla:", product_code)
                return None
            
            price = self.extract_price(soup)
            product_name = self.extract_product_name(soup)
            
            if product_code:
                print("tuotekoodi:", product_code)
            if price and product_name:
                print("tuotenimi:", product_name)
                print("hinta:", price, "€")
                return url, price, product_name
        else:
            print("Puuilo: Error with handling HTML.")
        return None

    def search_result_check(self, soup):
        count_span = soup.find('span', class_='amsearch-results-count')
        if count_span and count_span.get_text(strip=True) == '(0)':
            return True
        else:
            return False
    
    def extract_product_name(self, soup):
        product_name_element = soup.find('h2', class_='product name product-item-name')
        if product_name_element:
            product_name = product_name_element.find('a', class_='product-item-link').get_text(strip=True)
            return product_name
        else:
            return None

    def extract_price(self, soup):
        price_span = soup.find('span', class_='price')
        if price_span:
            price_text = price_span.get_text(strip=True).replace('\xa0', '')
            return float(price_text.replace("€", "").replace(",", "."))

        return None

class KRautaScraper(Scraper):
    def __init__(self):
        super().__init__()
    def search_product(self, product_code):
        url = f"https://www.k-rauta.fi/etsi?query={product_code}"
        html = self.get_html_from_url(url)

        if html:
            soup = BeautifulSoup(html, 'html.parser')
            price = self.extract_price(soup)
            product_name = self.extract_product_name(soup)
            formatted_product_name = product_name.replace(" ", "-").replace("/", "").lower()
            item_page_url = f"https://www.k-rauta.fi/tuote/{formatted_product_name}/{product_code}"
            brand_name = self.specific_product_page(item_page_url)
            print("\nK-RAUTA:")
            if product_code:
                print("tuotekoodi:", product_code)
            if price and product_name:
                print("tuotenimi:", product_name)
                print("hinta:", price, "€")
                return url, price, product_name, brand_name
        else:
            print("K-rauta: Error with handling HTML.")
        return None

    def extract_price(self, soup):
        price_span = soup.find('span', class_='price-view-fi__sale-price--prefix')
        if price_span:
            price_text = price_span.next_sibling.strip()
            return float(price_text.replace("€", "").replace(",", "."))

        return None

    def extract_product_name(self, soup):
        product_name_h2 = soup.find('h2', class_='product-card__name', attrs={'data-cy': 'product-card-name'})
        if product_name_h2:
            product_name = product_name_h2.get_text(strip=True)
            return product_name
        else:
            return None

    def specific_product_page(self, url):
        html = self.get_html_from_url(url)
        if html:
            soup = BeautifulSoup(html, 'html.parser')
            brand_name = soup.find('a', class_='product-heading__brand-name').text
            return brand_name
        return None

class TokmanniScraper(Scraper):
    def __init__(self):
        super().__init__()
    def search_product(self, product_code):
        url = f"https://www.tokmanni.fi/search/?q={product_code}"
        html = self.get_html_from_url(url, 10)

        if html:
            soup = BeautifulSoup(html, 'html.parser')
            print("\nTOKMANNI:")
            if self.no_results(soup):
                print("Ei tuloksia haulla:", product_code)
                return None

            price, product_name = self.extract_price_and_name(soup)

            if price and product_name:
                print("tuotekoodi:", product_code)
                print("tuotenimi:", product_name)
                print("hinta:", price, "€")
                return url, price, product_name
        else:
            print("Tokmanni: Error with handling HTML.")
        return None

    def no_results(self, soup):
        no_results = soup.find('div', class_='kuNoResults-lp-message')
        if no_results:
            return True
        else:
            return False

    def extract_price_and_name(self, soup):
        price_div = soup.find('div', class_='kuSalePrice')
        name_element = soup.select_one('.kuName a')
        product_name = name_element.get_text(strip=True)
        if price_div and product_name:
            whole_part = price_div.find('span', class_='ku-coins').previous_sibling
            decimal_part = price_div.find('span', class_='ku-coins').text
            price_str = whole_part.replace('\xa0', '') + decimal_part.replace('\xa0', '').replace('€', '')

            try:
                price_float = float(price_str) / 100
                return price_float, product_name
            except ValueError as e:
                print(f"Error converting price to float: {e}")
        else:
            return None


def format_product_link(url, product_name, sheet, given_row, given_column):
    cell = sheet.cell(row=given_row, column=given_column)
    cell.value = product_name
    cell.hyperlink = url
    return None

class ExcelHandler:
    def handle(self):
        file_path = "C:/Users/Karoliina/kilpailijahintoja.xlsx"
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        frozen_pane = sheet.sheet_view.pane
        frozen_pane_height = int(frozen_pane.ySplit) if frozen_pane.ySplit is not None else 0
        target_column_index = 1
        blank_cell_encountered = False

        with concurrent.futures.ThreadPoolExecutor() as executor:
            # Create instances of scraper classes directly
            prisma_scraper = PrismaScraper()
            k_rauta_scraper = KRautaScraper()
            puuilo_scraper = PuuiloScraper()
            tokmanni_scraper = TokmanniScraper()
            # List of functions to execute in parallel
            functions_to_execute = [
                prisma_scraper.search_product,
                k_rauta_scraper.search_product,
                puuilo_scraper.search_product,
                tokmanni_scraper.search_product,
            ]
            
            for row in sheet.iter_rows(min_row=frozen_pane_height + 1, max_row=sheet.max_row, min_col=1, max_col=12):
                if target_column_index >= len(row):
                    new_column = sheet.max_column + 1
                    sheet.cell(row=1, column=new_column, value=f'Column{new_column}')

                if row[0].value is not None:
                    # Execute functions in parallel
                    results = list(executor.map(lambda f: f(row[0].value), functions_to_execute))
                    # Process results and update the sheet accordingly
                    # My first time using lambda or concurrent.futures so allow it
                    prisma_data = []
                    k_rauta_data = []
                    puuilo_data = []
                    tokmanni_data = []
                    if results[0]:
                        for result in results[0]:
                            if result:
                                prisma_data.append(result)
                    else:
                        prisma_data = None

                    if results[1]:
                        for result in results[1]:
                            if result:
                                k_rauta_data.append(result)
                    else:
                        k_rauta_data = None

                    if results[2]:
                        for result in results[2]:
                            if result:
                                puuilo_data.append(result)
                    else:
                        puuilo_data = None
                    if results[3]:
                        for result in results[3]:
                            if result:
                                tokmanni_data.append(result)
                    else:
                        tokmanni_data = None
                            
                    if prisma_data:
                        sheet.cell(row=row[0].row, column=target_column_index + 1, value=prisma_data[0].upper())
                        sheet.cell(row=row[0].row, column=target_column_index + 2, value=prisma_data[2])

                    if k_rauta_data:
                        sheet.cell(row=row[0].row, column=target_column_index + 3, value=k_rauta_data[1])
                        current_row = row[0].row
                        current_column = target_column_index + 4
                        format_product_link(k_rauta_data[0], k_rauta_data[2], sheet, current_row, current_column)
                    else:
                        sheet.cell(row=row[0].row, column=target_column_index + 3, value="")
                        sheet.cell(row=row[0].row, column=target_column_index + 4, value="")

                    if puuilo_data:
                        sheet.cell(row=row[0].row, column=target_column_index + 7, value=puuilo_data[1])
                        current_row = row[0].row
                        current_column = target_column_index + 8
                        format_product_link(puuilo_data[0], puuilo_data[2], sheet, current_row, current_column)
                    else:
                        sheet.cell(row=row[0].row, column=target_column_index + 7, value="")
                        sheet.cell(row=row[0].row, column=target_column_index + 8, value="")

                    if tokmanni_data:
                        sheet.cell(row=row[0].row, column=target_column_index + 9, value=tokmanni_data[1])
                        current_row = row[0].row
                        current_column = target_column_index + 10
                        format_product_link(tokmanni_data[0], tokmanni_data[2], sheet, current_row, current_column)
                    else:
                        sheet.cell(row=row[0].row, column=target_column_index + 9, value="")
                        sheet.cell(row=row[0].row, column=target_column_index + 10, value="")
                else:
                    blank_cell_encountered = True

                if blank_cell_encountered:
                    break

        wb.save(file_path)
        return 0

# Main
start_time = time.time()
handler = ExcelHandler()
handler.handle()
end_time = time.time()
elapsed_time = end_time - start_time
print(f"This program took: {elapsed_time:.2f} seconds to run. Kiitos ja anteeksi.")
