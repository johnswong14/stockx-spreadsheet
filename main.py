import config
import glob
import os
import time
from load_excel import load_excel_workbook
from load_excel import load_excel_spreadsheet
from load_excel import find_end_row
from update import update_price

# main program
def main():
    # if output folder does not exist, create directory
    if not os.path.isdir(config.OUTPUT_DIR):
        os.mkdir(config.OUTPUT_DIR)
    
    # check if output folder contains a spreadsheet
    list_of_files = glob.glob(os.path.join(config.OUTPUT_DIR, '*'))
    
    if len(list_of_files) != 0:
        list_of_files = glob.glob(f"{config.OUTPUT_DIR}/[!~]*.xlsx")
        input_filename = max(list_of_files, key=os.path.getctime)
        input_filename_without_path = os.path.basename(input_filename)
        print(f"Loading the lastest spreadsheet: {input_filename_without_path.lower()}")
    # prompt user to enter the filename of initial spreadsheet
    else:
        input_filename = input("Enter the filename of your spreadsheet (i.e. Template.xlsx): ")
        while not os.path.isfile(input_filename):
            print(f"{input_filename} does not exist")
            input_filename = input("Enter the filename of your spreadsheet (i.e. Template.xlsx): ")
        print(f"Loading the spreadsheet: {input_filename.lower()}")
    
    # load workbook then spreadsheet
    workbook = load_excel_workbook(input_filename)
    sheet = load_excel_spreadsheet(workbook)
    
    # find end row in spreadsheet
    end_row = find_end_row(sheet)

    if end_row:
        # update prices
        start = time.time()
        print("Updating StockX prices...\n")
        
        if update_price(sheet, end_row):
            output_filename = f"output_{int(time.time())}.xlsx"
            os.chdir(config.OUTPUT_DIR)
            workbook.save(filename=output_filename)
            end = time.time()
            duration = format(end - start, '.2f')
            print(f"Finished updating StockX prices in {duration}s")
    
# run main program
main()