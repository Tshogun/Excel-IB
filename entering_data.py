from loading_book import *

workbook = loading_file('./Tata_motors.xlsx')  
company_name, meta_data_dict, profit_and_loss_data_dict, quarter_data_dict, balance_sheet_dict, cash_flow_price_derived_dict = get_data_from_data_sheet(workbook)

def create_sheet__HistoricalFS(workbook):
    try:
        sheet_historicalFS = workbook.create_sheet("HistoricalFS", 0)
        print('[INFO] HistoricalFS sheet created')
        sheet_historicalFS = workbook.active
        print('[INFO] HistoricalFS sheet is activated')
        return sheet_historicalFS
    except (RuntimeError, TypeError, NameError):
        print(f'[INFO] Loading failed.')
        pass
    
sheet_historicalFS = create_sheet__HistoricalFS(workbook)
    
