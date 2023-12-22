import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
import concurrent.futures
from concurrent.futures import ThreadPoolExecutor
from openpyxl.utils import get_column_letter
import winsound

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
            if self.no_search_result_check(soup):
                #print("\nPRISMA:\nEi tuloksia haulla:", product_code)
                return None
            price = self.extract_price(soup)
            product_name = self.extract_product_name(soup)
            brand_name = product_name.split(" ")[0].upper()
            link = self.specific_product_page(soup)
            link = f"https://www.prisma.fi{link}"
            if product_name and price:
                #print("\nPRISMA:\ntuotekoodi:", product_code,"\ntuotenimi:", product_name,"\nhinta:", price, "€")
                return product_name, link, brand_name
        else:
            print("PRISMA: Error with handling HTML.")
        return None, None, None

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
                price_value = price_span.get_text(strip=True).replace("€","")
                return price_value
            else:
                return None
        else:
            return None

    def specific_product_page(self, soup):
        link_tag = soup.find('a', class_='ProductCard_blockLink__NTUGB') # Find the <a> tag with the specified class
        if link_tag:
            link = link_tag['href'] # Extract the value of the href attribute
            return link
        else:
            return None

class PuuiloScraper(Scraper):
    def __init__(self):
        super().__init__()
    def search_product(self, product_code):
        url = f"https://www.puuilo.fi/catalogsearch/result/?q={product_code}"
        html = self.get_html_from_url(url, 10)
        if html:
            soup = BeautifulSoup(html, 'html.parser')
            no_results = self.search_result_check(soup)
            if no_results:
                #print("\nPUUILO:\nEi tuloksia haulla:", product_code)
                return None
            price = self.extract_price(soup)
            product_name = self.extract_product_name(soup)
            link = self.specific_product_page(soup)
            if price and product_name:
                #print("\nPUUILO:\ntuotekoodi:", product_code, "\ntuotenimi:", product_name, "\nhinta:", price, "€")
                return link, price, product_name
        else:
            print("PUUILO: Error with handling HTML.")
        return None, None, None

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
        else:
            return None

    def specific_product_page(self, soup):
        a_element = soup.select_one('article.product-item-info a.product-item-photo') # Find the <a> element inside the <article> tag
        if a_element:
            link = a_element['href'] # Extract the value of the "href" attribute
            return link
        else:
            return None


class KRautaScraper(Scraper):
    def __init__(self):
        super().__init__()
    def search_product(self, product_code):
        url = f"https://www.k-rauta.fi/etsi?query={product_code}"
        html = self.get_html_from_url(url)

        if html:
            soup = BeautifulSoup(html, 'html.parser')
            # ADD FUNCTION TO CHECK IF PRODUCT WAS FOUND
            price = self.extract_price(soup)
            product_name = self.extract_product_name(soup)
            formatted_product_name = product_name.replace(" ", "-").replace("/", "").lower()
            link = self.specific_product_page(soup)
            link = f"https://www.k-rauta.fi{link}"
            if price and product_name:
                #print("\nK-RAUTA:\ntuotekoodi:", product_code, "\ntuotenimi:", product_name, "\nhinta:", price, "€")
                return link, price, product_name
        else:
            print("K-RAUTA: Error with handling HTML.")
        return None, None, None, None

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

    def specific_product_page(self, soup):
        a_element = soup.find('a', class_='product-card') # Find the <a> element with class "product-card"
        if a_element:
            link = a_element['href'] # Extract the value of the "href" attribute
            return link
        else:
            return None

class TokmanniScraper(Scraper):
    def __init__(self):
        super().__init__()
    def search_product(self, product_code):
        url = f"https://www.tokmanni.fi/search/?q={product_code}"
        html = self.get_html_from_url(url, 10)
        if html:
            soup = BeautifulSoup(html, 'html.parser')
            if self.no_results(soup):
                #print("\nTOKMANNI:\nEi tuloksia haulla:", product_code)
                return None

            price, product_name = self.extract_price_and_name(soup)
            link = self.specific_product_page(soup)
            if price and product_name:
                #print("\nTOKMANNI:\ntuotekoodi:", product_code, "\ntuotenimi:", product_name, "\nhinta:", price, "€")
                return link, price, product_name
        else:
            print("TOKMANNI: Error with handling HTML.")
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
                return None, None
        else:
            return None, None

    def specific_product_page(self, soup):
        a_element = soup.select_one('div.klevuImgWrap a') # Find the <a> element inside the <div> with class "klevuImgWrap"
        if a_element:
            link = a_element['href'] # Extract the value of the "href" attribute
            return link
        else:
            return None

class BauhausScraper(Scraper):
    def __init__(self):
        super().__init__()
    def search_product(self, product_code):
        query = f"bauhaus {product_code}"
        url = f"https://www.google.com/search?q={query}"
        html = self.get_html_from_url(url)
        if html:
            soup = BeautifulSoup(html, 'html.parser')
            link = self.search_result_check(soup, product_code)
            if not link:
                #print("\nBAUHAUS:\nEi tuloksia haulla:", product_code)
                return None
            html = self.get_html_from_url(link)
            if html:
                soup = BeautifulSoup(html, 'html.parser')
                price = self.extract_price(soup)
                product_name = self.extract_product_name(soup)
                if product_name and price:
                    #print("\nBAUHAUS:\ntuotekoodi:", product_code,"\ntuotenimi:", product_name,"\nhinta:", price, "€")
                    return link, price, product_name
            else:
                print("BAUHAUS: Error with handling product page HTML.")
        else:
            print("BAUHAUS: Error with handling Google search HTML.")
        return None, None

    # Return False or link to the product page in Bauhaus
    def search_result_check(self, soup, product_code):
        emphasized_text_tag = soup.find('em')  # Find the first <em> tag
        if not emphasized_text_tag: # No results based on the EAN
            return False
        emphasized_text = emphasized_text_tag.get_text(strip=True)       
        target_span = soup.find('span', {'class': 'VuuXrf'})  # Find the name of the website in Google search result
        target_span_parent_div = ""
        if target_span:
            target_span_parent_div = target_span.find_parent('div') # Find the parent div and get its class attribute
            if target_span_parent_div:
                target_span_parent_div = target_span_parent_div.get('class', '')
        if str(emphasized_text) == str(product_code) and str(target_span.text) == "bauhaus.fi":  # Check if the emphasized text is the EAN code and the website is Bauhaus
            div_element = soup.find('div', class_='kb0PBd')
            link_element = div_element.find('a') # Find the <a> element and extract the link
            if not link_element:
                return False
            link = link_element.get('href')
            return link
        else: # emphasized text was not the EAN code or website was not bauhaus
            return False

    def extract_product_name(self, soup):
        name_span = soup.find('h1', class_='page-title').find('span', class_='base') # Find the <span> element with class "base" and itemprop "name"
        if name_span:
            name_text = name_span.get_text(strip=True) # Extract the text content from the name_span
            return name_text
        else:
            return None

    def extract_price(self, soup):
        price_span = soup.find('span', class_='price') # Find the <span> element with class "price"
        if price_span:
            price_text = price_span.get_text(strip=True) # Extract the text content from the price_span
            price = float(price_text.replace('\xa0', ''))/100
            return price
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
            bauhaus_scraper = BauhausScraper()
            # List of functions to execute in parallel
            functions_to_execute = [
                prisma_scraper.search_product,
                k_rauta_scraper.search_product,
                puuilo_scraper.search_product,
                tokmanni_scraper.search_product,
                bauhaus_scraper.search_product,
            ]
            
            for row in sheet.iter_rows(min_row=frozen_pane_height + 1, max_row=sheet.max_row, min_col=1, max_col=12):
                if target_column_index >= len(row):
                    new_column = sheet.max_column + 1
                    sheet.cell(row=1, column=new_column, value=f'Column{new_column}')

                if row[0].value is not None:
                    # Execute functions in parallel
                    #|        PRISMA     |      K-RAUTA     |       PUUILO     |      TOKMANNI    |      BAUHAUS     |
                    #|   [0]   [1]   [2] | [0]  [1]    [2]  | [0]  [1]    [2]  | [0]  [1]   [2]   | [0]  [1]    [2]  |
                    #|PRODUCT|LINK |BRAND|LINK|PRICE|PRODUCT|LINK|PRICE|PRODUCT|LINK|PRICE|PRODUCT|LINK|PRICE|PRODUCT|
                    results = list(executor.map(lambda f: f(row[0].value), functions_to_execute))
                    # Process results and update the sheet accordingly
                    # My first time using lambda or concurrent.futures so allow it
                    if results[0]: #PRISMA INFO TO EXCEL
                        sheet.cell(row=row[0].row, column=target_column_index + 1, value=results[0][0].upper()) # product name
                        sheet.cell(row=row[0].row, column=target_column_index + 2, value=results[0][2])         # brand
                        current_row = row[0].row
                        current_column = target_column_index + 1
                        format_product_link(results[0][1], results[0][0], sheet, current_row, current_column)   # link with product name as label
                    if results[1]: #K-RAUTA INFO TO EXCEL
                        sheet.cell(row=row[0].row, column=target_column_index + 3, value=results[1][1])         # price
                        current_row = row[0].row
                        current_column = target_column_index + 4
                        format_product_link(results[1][0], results[1][2], sheet, current_row, current_column)   # link with product name as label
                    else:
                        sheet.cell(row=row[0].row, column=target_column_index + 3, value="")                    # leave K-rauta price empty
                        sheet.cell(row=row[0].row, column=target_column_index + 4, value="")                    # leave K-rauta link empty
                    if results[4]: #BAUHAUS INFO TO EXCEL
                        sheet.cell(row=row[0].row, column=target_column_index + 5, value=results[4][1])
                        current_row = row[0].row
                        current_column = target_column_index + 6
                        format_product_link(results[4][0], results[4][2], sheet, current_row, current_column)
                    else:
                        sheet.cell(row=row[0].row, column=target_column_index + 5, value="")
                        sheet.cell(row=row[0].row, column=target_column_index + 6, value="")
                    if results[2]: #PUUILO INFO TO EXCEL
                        sheet.cell(row=row[0].row, column=target_column_index + 7, value=results[2][1])
                        current_row = row[0].row
                        current_column = target_column_index + 8
                        format_product_link(results[2][0], results[2][2], sheet, current_row, current_column)
                    else:
                        sheet.cell(row=row[0].row, column=target_column_index + 7, value="")
                        sheet.cell(row=row[0].row, column=target_column_index + 8, value="")

                    if results[3]: #TOKMANNI INFO TO EXCEL
                        sheet.cell(row=row[0].row, column=target_column_index + 9, value=results[3][1])
                        current_row = row[0].row
                        current_column = target_column_index + 10
                        format_product_link(results[3][0], results[3][2], sheet, current_row, current_column)
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
# Play a sound when done
if __name__ == "__main__":
    print(f"This program took: {elapsed_time:.2f} seconds to run. Kiitos ja anteeksi.")
    for _ in range(3):
        winsound.Beep(1000, 500)
