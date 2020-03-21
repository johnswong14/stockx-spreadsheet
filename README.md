# stockx-spreadsheet

StockX is an online marketplace known for primarily reselling sought after sneakers and clothes.\
This script will scrape StockX for the market price of a particular item and update the price in a spreadsheet.\
Template is included and must be used in order for this application to be successful :)

## Requirements

-   Python 3
-   pip

## How to Run

1. Install Python 3: https://www.python.org/downloads
2. Install pip: https://pip.pypa.io/en/stable/installing
3. Open Terminal (Mac) / Windows (CMD)
4. Navigate to directory containing source code
5. Type: pip3 install -r requirements.txt
6. Once step 5 is finished, type: python3 main.py
7. Follow the prompts shown in command line

## Notes

-   Make sure you've installed everything; refer to the requirements
-   Don't forget to input your items into the spreadsheet
-   Only fill in 'Item', 'Size', and 'Price Paid' fields
-   For better results, enter the item name as it appears on StockX
-   If size is not applicable, leave it blank
-   Don't change the template of the spreadsheet! This is critical for the script to run smoothly!
-   If you run into repeated timed out errors, try increasing the update delay (in config.py)
-   Spreadsheet with updated prices will be in the output folder
