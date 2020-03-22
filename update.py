import config
import time
from load_excel import is_format_valid
from product import get_ticker_symbol
from product import get_prices

# update StockX price
def update_price(sheet, end_row):
    total_items = 0
    total_items_price_up = 0
    total_items_price_down = 0
    
    try:
        if is_format_valid(sheet):
            for value in sheet.iter_rows(min_row=138,
                                    max_row=end_row,
                                    min_col=1,
                                    max_col=10,
                                    values_only=False):
                product_info = None
                product = {
                    "title": value[0].value,
                    "size": value[1].value,
                    "ticker_symbol": value[2].value
                }
                
                if product['title']:
                    if product['ticker_symbol']:
                        product_info = get_prices(product)
                    else:
                        ticker_symbol, search_result_products = get_ticker_symbol(product)
                        if (ticker_symbol):
                            value[2].value = ticker_symbol
                            product['ticker_symbol'] = value[2].value
                            product_info = get_prices(product, search_result_products)
                    
                    if product_info:
                        total_items += 1
                        previous_price = float(value[6].value) if value[6].value else 0
                        print(f"Previous price: {format(previous_price, '.2f')}")
                        
                        if product_info['lowest_ask_price']:
                            value[6].value = product_info['lowest_ask_price']
                            print(f"Lowest Ask Price: {format(product_info['lowest_ask_price'], '.2f')}")
                        elif product_info['last_sale_price']:
                            value[6].value = product_info['last_sale_price']
                            print(f"Last Sale Price: {format(product_info['last_sale_price'], '.2f')}")
                        elif product_info['highest_bid_price']:
                            value[6].value = product_info['highest_bid_price']
                            print(f"Highest Bid Price: {format(product_info['highest_bid_price'], '.2f')}")
                            
                        if value[6].value > previous_price:
                            total_items_price_up += 1
                            print("PRICE IS BOOMING!\n")
                        elif value[6].value < previous_price:
                            total_items_price_down += 1
                            print("Price went down...\n")
                        else:
                            print("Price stayed the same.\n")
                    
                    time.sleep(config.UPDATE_DELAY)
                else:
                    continue
            
            total_items_price_same = total_items - (total_items_price_up + total_items_price_down)
            
            print(f"Total items: {total_items}")
            print(f"Total items that went up in price: {total_items_price_up}")
            print(f"Total items that went down in price: {total_items_price_down}")
            print(f"Total items that stayed the same in price: {total_items_price_same}\n")
            return True
        else:
            print("SPREADSHEET TEMPLATE IS INVALID. PLEASE USE THE TEMPLATE THAT WAS GIVEN.")
            return False
    except KeyboardInterrupt:
        print("Stopped updating prices")
        return False
    except Exception as e:
        print(f"Error updating prices:")
        print(f"{e}")
        return False