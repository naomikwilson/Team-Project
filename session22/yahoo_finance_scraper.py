import csv
from bs4 import BeautifulSoup
import requests

DOWNLOAD_URL = 'https://finance.yahoo.com/trending-tickers'


def download_page(url):
    """
    Reference: https://www.scrapehero.com/scrape-yahoo-finance-stock-market-data/
    """
    headers = {
        "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
        "accept-encoding": "gzip, deflate, br",
        "accept-language": "en-GB,en;q=0.9,en-US;q=0.8,ml;q=0.7",
        "cache-control": "max-age=0",
        "dnt": "1",
        "sec-fetch-dest": "document",
        "sec-fetch-mode": "navigate",
        "sec-fetch-site": "none",
        "sec-fetch-user": "?1",
        "upgrade-insecure-requests": "1",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.122 Safari/537.36",
    }
    response = requests.get(url, headers=headers, timeout=30)
    return response.text


# print(download_page(DOWNLOAD_URL))


def parse_html(html):
    """
    Analyze the html page, find the information and return the list of tuples (stock_name, symbol, price, change, change_percent, volume, market_cap)
    """
    soup = BeautifulSoup(html, 'html.parser')
    # print(soup.prettify())
    stock_table = soup.find('tbody')
    stock_list = []
    # print(stock_table.prettify())
    for stock_row in stock_table.find_all('tr'):
        stock_name = stock_row.find('td', attrs={'aria-label': 'Name'}).string
        # print(stock_name)

        symbol_element = stock_row.find('td', attrs={'aria-label': 'Symbol'})
        symbol = symbol_element.find('a', attrs={'class': 'Fw(600) C($linkColor)'}).string
        # print(symbol)

        price_element = stock_row.find('td', attrs={'aria-label': 'Last Price'})
        price = price_element.find('fin-streamer', attrs={'class': ''}).string
        # print(price)

        change_element_outer = stock_row.find('td', attrs={'aria-label': 'Change'})
        change_element_inner = change_element_outer.find('fin-streamer', attrs={'class': 'Fw(600)'})
        change = change_element_inner.find('span').string
        # print(change)

        change_percent_element_outer = stock_row.find('td', attrs={'aria-label': '% Change'})
        change_percent_element_inner = change_percent_element_outer.find('fin-streamer', attrs={'class': 'Fw(600)'})
        change_percent = change_percent_element_inner.find('span').string
        # print(change_percent)

        volume_element = stock_row.find('td', attrs={'aria-label': 'Volume'})
        volume = volume_element.find('fin-streamer', attrs={'class': ''}).string
        # print(volume)

        market_cap_element = stock_row.find('td', attrs={'aria-label': 'Market Cap'})
        try:
            market_cap = market_cap_element.find('fin-streamer', attrs={'class': ''}).string
        except: # market caps with "N/A" values have the element "<span>N/A</span>", so the above code line would return an error
            market_cap = "N/A"
        # print(market_cap)

        stock_list.append((stock_name, symbol, price, change, change_percent, volume, market_cap))
    return stock_list


# parse_html(download_page(DOWNLOAD_URL))
# print(len(parse_html(download_page(DOWNLOAD_URL))))

def csv_conversion(url):
    """
    Convert information in Yahoo Finance's Trending Tickers page into a CSV.
    """
    with open('data/yahoo_trending_tickers.csv', 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f)

        fields = ('stock_name', 'symbol', 'price', 'change', 'change_percent', 'volume', 'market_cap')
        writer.writerow(fields)

        html = download_page(url)
        stocks = parse_html(html)
        writer.writerows(stocks)

def main():
    url = DOWNLOAD_URL
    csv_conversion(url)



if __name__ == '__main__':
    main()