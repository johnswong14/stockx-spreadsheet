import config
import requests
import json

# retrieve ticker symbol of requested item
def get_ticker_symbol(requested_item):
    search_request_payload = {
        'params': 'query={0}&hitsPerPage=20&facets=*'.format(requested_item['title'])
    }
    ticker_symbol = None
    search_result_products = None
    
    try:
        search_result = requests.post(config.STOCKX_SEARCH_ENDPOINT, headers=config.REQUEST_HEADERS, params=config.SEARCH_REQUEST_PARAMS, json=search_request_payload).json()['hits']
        
        if search_result:
            item = search_result[0]
            ticker_symbol = item['ticker_symbol']
        else:
            print(f"{requested_item['title']} could not be found on StockX.")
            print("Please check and make sure the item name matches the item found on StockX.\n")
    except Exception as e:
        print(f"Failed retrieving ticker symbol for {requested_item['title']}:")
        print(f"{e}\n")
        
    return ticker_symbol, search_result_products

# retrieve prices of requested item
def get_prices(requested_item, search_result_products=None):
    if requested_item['ticker_symbol']:
        keyword = requested_item['title']
        URL = f"{config.STOCKX_PRODUCT_ENDPOINT}{keyword}"
        product_info = None
        
        if not search_result_products:
            try:
                search_result = requests.get(URL, headers=config.REQUEST_HEADERS)
                if search_result.status_code == 200:
                    search_result = search_result.json()
                    search_result_products = search_result['Products']
                else:
                    print(f"Failed retrieving information for {requested_item['title']}:")
                    return
            except Exception as e:
                print(f"Failed retrieving information for {requested_item['title']}:")
                print(f"{e}\n")
                return
        
        for product in search_result_products:
            if (product['tickerSymbol'] == requested_item['ticker_symbol']
                    and product['shoeSize'] == requested_item['size']):
                product_info = {
                    "title": product['title'],
                    "actual_size": product['shoeSize'], 
                    "lowest_ask_price": product['market']['lowestAsk'], 
                    "highest_bid_price": product['market']['highestBid'], 
                    "last_sale_price": product['market']['lastSale']
                }

                print(f"Product: {product_info['title']}\nSize: {product_info['actual_size']}")
                break
            
        if product_info:
            return product_info
        else:
            print(f"{requested_item['title']} could not be found on StockX.")
            print("Please check and make sure the item name, ticker symbol, and size matches the item found on StockX.\n")
    else:
        print("Missing StockX ticker symbol\n")