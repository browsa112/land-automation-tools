import pandas as pd
import json
import glob
import openpyxl
from openpyxl import Workbook

all_data = {}
counties_states = []

'''
State 
County
Acreage Range

For Sale .25 to 1 ac	Sold .25 to 1 ac	.25 to 1 acre ratio
For Sale 1 to 2 ac	Sold 1 to 2 ac	1 to 2 acre ratio
For Sale 2 to 5 ac	Sold 2 to 5 ac	2 to 5 acres ratio
For Sale 5 to 10 ac	Sold 5 to 10 ac	5 to 10 acres ratio
For Sale 10 to 20 ac	Sold 10 to 20 ac	10 to 20 acres ratio
For Sale 20 to 50 ac	Sold 20 to 50 ac	20 to 50 acres ratio
For Sale 50 to max	Sold 50 to max	50 to max acres ratio
'''

ACRE_SIZE = 43560 
IS_FORMAT_2 = False
LOT_SIZE = "lotSize"
SALE_TYPE = "statusType"
PRICE = "price"

def calculate_acres(row):
   if row['Lot Area Unit'] == "acres":
      return row['Lot Area']
   if row['Lot Area Unit'] == "sqft":
      return row['Lot Area'] / ACRE_SIZE
   return "other"

def drop_missing_units(df):
    print(df.head())
    df.dropna(subset = ['Lot Area Unit'], inplace=True)

def get_acres(lot_size):
    return lot_size / 43560

def drop_extra_rows(df_sold, df_on_market):
    df_sold.drop(df_sold.loc[df_sold[SALE_TYPE]=="In accordance with local MLS rules, some MLS listings are not included in the download"].index, inplace=True)
    df_on_market.drop(df_on_market.loc[df_on_market[SALE_TYPE]=="In accordance with local MLS rules, some MLS listings are not included in the download"].index, inplace=True)

def check_format(df_on_market):
    global LOT_SIZE
    global SALE_TYPE
    global PRICE
    global IS_FORMAT_2

    if "Lot Area" in df_on_market:
        LOT_SIZE = "Lot Area"
        SALE_TYPE = "Status Type"
        PRICE = "Sold Price"
        IS_FORMAT_2 = True
    else:
        LOT_SIZE = 'LOT SIZE'
        SALE_TYPE = 'SALE TYPE'
        PRICE = "PRICE"
        IS_FORMAT_2 = False

def go_through_rows():
    global LOT_SIZE
    global SALE_TYPE
    global PRICE
    global IS_FORMAT_2

    for file in glob.glob("*sold.csv"):
        df_sold = pd.read_csv(file)
        df_on_market = pd.read_csv(file.replace("sold", "sale"))

        check_format(df_on_market)
        drop_extra_rows(df_sold, df_on_market)

        if IS_FORMAT_2 == False:
            df_point_25_1 = df_sold[(df_sold[LOT_SIZE] >= (ACRE_SIZE * .25)) & (df_sold[LOT_SIZE] <= (ACRE_SIZE * 1))]
            df_1_2 = df_sold[(df_sold[LOT_SIZE] >= (ACRE_SIZE * 1)) & (df_sold[LOT_SIZE] <= (ACRE_SIZE * 2))]
            df_2_5 = df_sold[(df_sold[LOT_SIZE] >= (ACRE_SIZE * 2)) & (df_sold[LOT_SIZE] <= (ACRE_SIZE * 5))]
    
            df_point_25_1_sales = df_on_market[(df_on_market[LOT_SIZE] >= (ACRE_SIZE * .25)) & (df_on_market[LOT_SIZE] <= (ACRE_SIZE * 1))]
            df_1_2_sales = df_on_market[(df_on_market[LOT_SIZE] >= (ACRE_SIZE * 1)) & (df_on_market[LOT_SIZE] <= (ACRE_SIZE * 2))]
            df_2_5_sales = df_on_market[(df_on_market[LOT_SIZE] >= (ACRE_SIZE * 2)) & (df_on_market[LOT_SIZE] <= (ACRE_SIZE * 5))]


            df_5_10 = df_sold[(df_sold[LOT_SIZE] >= (ACRE_SIZE * 5)) & (df_sold[LOT_SIZE] <= (ACRE_SIZE * 10))]
            df_10_20 = df_sold[(df_sold[LOT_SIZE] >= (ACRE_SIZE * 10)) & (df_sold[LOT_SIZE] <= (ACRE_SIZE * 20))]
            df_20_50 = df_sold[(df_sold[LOT_SIZE] >= (ACRE_SIZE * 20)) & (df_sold[LOT_SIZE] <= (ACRE_SIZE * 50))]
            df_50_plus = df_sold[df_sold[LOT_SIZE] >= (ACRE_SIZE * 50)]
    
            df_5_10_sales = df_on_market[(df_on_market[LOT_SIZE] >= (ACRE_SIZE * 5)) & (df_on_market[LOT_SIZE] <= (ACRE_SIZE * 10))]
            df_10_20_sales = df_on_market[(df_on_market[LOT_SIZE] >= (ACRE_SIZE * 10)) & (df_on_market[LOT_SIZE] <= (ACRE_SIZE * 20))]
            df_20_50_sales = df_on_market[(df_on_market[LOT_SIZE] >= (ACRE_SIZE * 20)) & (df_on_market[LOT_SIZE] <= (ACRE_SIZE * 50))]
            df_50_plus_sales = df_on_market[df_on_market[LOT_SIZE] >= (ACRE_SIZE * 50)]
        else:

            df_sold.head()
            df_on_market.head()
            drop_missing_units(df_sold)
            drop_missing_units(df_on_market)

            df_sold['ACRES'] = df_sold.apply(calculate_acres, axis=1)
            df_on_market['ACRES'] = df_on_market.apply(calculate_acres, axis=1)

            df_point_25_1 = df_sold[(df_sold['ACRES'] >= (.25)) & (df_sold['ACRES']<= (1))]
            df_1_2 = df_sold[(df_sold['ACRES'] >= (1)) & (df_sold['ACRES'] <= (2))]
            df_2_5 = df_sold[(df_sold['ACRES'] >= (2)) & (df_sold['ACRES'] <= (5))]
    
            df_point_25_1_sales = df_on_market[(df_on_market['ACRES'] >= (.25)) & (df_on_market['ACRES'] <= (1))]
            df_1_2_sales = df_on_market[(df_on_market['ACRES'] >= (1)) & (df_on_market['ACRES'] <= (2))]
            df_2_5_sales = df_on_market[(df_on_market['ACRES'] >= (2)) & (df_on_market['ACRES'] <= (5))]


            df_5_10 = df_sold[(df_sold['ACRES'] >= (5)) & (df_sold['ACRES'] <= (10))]
            df_10_20 = df_sold[(df_sold['ACRES'] >= (10)) & (df_sold['ACRES'] <= (20))]
            df_20_50 = df_sold[(df_sold['ACRES'] >= (20)) & (df_sold['ACRES'] <= (50))]
            df_50_plus = df_sold[df_sold['ACRES'] >= (50)]
    
            df_5_10_sales = df_on_market[(df_on_market['ACRES'] >= (5)) & (df_on_market['ACRES'] <= (10))]
            df_10_20_sales = df_on_market[(df_on_market['ACRES'] >= (10)) & (df_on_market['ACRES'] <= (20))]
            df_20_50_sales = df_on_market[(df_on_market['ACRES'] >= (20)) & (df_on_market['ACRES'] <= (50))]
            df_50_plus_sales = df_on_market[df_on_market['ACRES'] >= (50)]
    
        # split data frame in to 20, 50 and 50+ acres    
        file_arr = file.split("_")
        county = file_arr[0]
        state = file_arr[1]
    
        if state not in all_data:
            all_data[state] = {}
        if county not in all_data[state]:
            all_data[state][county] = {".25-1": {"sold_to_sale_ratio": "", "solds": "", "on-market": ""}, "1-2": {"sold_to_sale_ratio": "", "solds": "", "on-market": ""}, "2-5": {"sold_to_sale_ratio": "", "solds": "", "on-market": ""}, "5-10": {"sold_to_sale_ratio": "", "solds": "", "on-market": ""}, "10-20":  { "sold_to_sale_ratio": "", "solds": "", "on-market": ""}, "20-50": {"sold_to_sale_ratio": "", "solds": "", "on-market": ""}, "50+": { "sold_to_sale_ratio": "", "solds": "", "on-market": ""}}
    
        if IS_FORMAT_2 == False:
            df_point_25_1['ACRES'] = df_point_25_1[LOT_SIZE].apply(get_acres)
            df_1_2['ACRES'] = df_1_2[LOT_SIZE].apply(get_acres)
            df_2_5['ACRES'] = df_2_5[LOT_SIZE].apply(get_acres)

            df_5_10['ACRES'] = df_5_10[LOT_SIZE].apply(get_acres)
            df_10_20['ACRES'] = df_10_20[LOT_SIZE].apply(get_acres)
            df_20_50['ACRES'] = df_20_50[LOT_SIZE].apply(get_acres)
            df_50_plus['ACRES'] = df_50_plus[LOT_SIZE].apply(get_acres)

        if df_point_25_1_sales.shape[0] > 0:
            all_data[state][county][".25-1"]["sold_to_sale_ratio"] = df_point_25_1.shape[0] / df_point_25_1_sales.shape[0]
        else:
            all_data[state][county][".25-1"]["sold_to_sale_ratio"] = "N/A"

        if df_1_2_sales.shape[0] > 0:
            all_data[state][county]["1-2"]["sold_to_sale_ratio"] = df_1_2.shape[0] / df_1_2_sales.shape[0]
        else:
            all_data[state][county]["1-2"]["sold_to_sale_ratio"] = "N/A"

        if df_2_5_sales.shape[0] > 0:
            all_data[state][county]["2-5"]["sold_to_sale_ratio"] = df_2_5.shape[0] / df_2_5_sales.shape[0]
        else:
            all_data[state][county]["2-5"]["sold_to_sale_ratio"] = "N/A"
        
        if df_5_10_sales.shape[0] > 0:
            all_data[state][county]["5-10"]["sold_to_sale_ratio"] = df_5_10.shape[0] / df_5_10_sales.shape[0]
        else:
            all_data[state][county]["5-10"]["sold_to_sale_ratio"] = "N/A"
    
        if df_10_20_sales.shape[0] > 0:
            all_data[state][county]["10-20"]["sold_to_sale_ratio"] = df_10_20.shape[0] / df_10_20_sales.shape[0]
        else:
            all_data[state][county]["10-20"]["sold_to_sale_ratio"] = "N/A"

        if df_20_50_sales.shape[0] > 0:
            all_data[state][county]["20-50"]["sold_to_sale_ratio"] = df_20_50.shape[0] / df_20_50_sales.shape[0]
        else:
            all_data[state][county]["20-50"]["sold_to_sale_ratio"] = "N/A"
        
        if df_50_plus_sales.shape[0] > 0:
            all_data[state][county]["50+"]["sold_to_sale_ratio"] = df_50_plus.shape[0] / df_50_plus_sales.shape[0]
        else:
            all_data[state][county]["50+"]["sold_to_sale_ratio"] = "N/A"

        all_data[state][county][".25-1"]["solds"] = df_point_25_1.shape[0]
        all_data[state][county]["1-2"]["solds"] = df_1_2.shape[0]
        all_data[state][county]["2-5"]["solds"] = df_2_5.shape[0] 

        all_data[state][county]["5-10"]["solds"] = df_5_10.shape[0]
        all_data[state][county]["10-20"]["solds"] = df_10_20.shape[0]
        all_data[state][county]["20-50"]["solds"] = df_20_50.shape[0] 
        all_data[state][county]["50+"]["solds"] = df_50_plus.shape[0]

        all_data[state][county][".25-1"]["on-market"] = df_point_25_1_sales.shape[0]
        all_data[state][county]["1-2"]["on-market"] = df_1_2_sales.shape[0]
        all_data[state][county]["2-5"]["on-market"] = df_2_5_sales.shape[0] 
    
        all_data[state][county]["5-10"]["on-market"] = df_5_10_sales.shape[0]
        all_data[state][county]["10-20"]["on-market"] = df_10_20_sales.shape[0]
        all_data[state][county]["20-50"]["on-market"] = df_20_50_sales.shape[0] 
        all_data[state][county]["50+"]["on-market"] = df_50_plus_sales.shape[0]
    
go_through_rows()
json_formatted_str = json.dumps(all_data, indent=2) 
print(json_formatted_str)

wb = openpyxl.load_workbook("mailing_log.xlsx") 
ws = wb.active

array_of_data = []

for state in all_data:
    for county in all_data[state]:
        county_val = [state, county, all_data[state][county]]
        counties_states.append(county_val)

row_pos = 0
iter_pos = 0
for row in ws.rows:

    if row[1].value == "Sold divided by For Sale" or row[1] == "Solds --> Last 12 months" or row[1] == "For Sale --> no time constraint":
        continue

    if row_pos > len(counties_states) - 1:
        break

    iter_pos = iter_pos + 1

    if iter_pos < 3:
        continue

    row[0].value = counties_states[row_pos][0]
    row[1].value = counties_states[row_pos][1]

    # For Sale .25 to 1 ac	Sold .25 to 1 ac
    row[5].value = counties_states[row_pos][2][".25-1"]["on-market"]
    row[6].value = counties_states[row_pos][2][".25-1"]["solds"]

    # For Sale 1 to 2 ac	Sold 1 to 2 ac
    row[14].value = counties_states[row_pos][2]["1-2"]["on-market"]
    row[15].value = counties_states[row_pos][2]["1-2"]["solds"]
    # For Sale 2 to 5 ac	Sold 2 to 5 ac
    row[23].value = counties_states[row_pos][2]["2-5"]["on-market"]
    row[24].value = counties_states[row_pos][2]["2-5"]["solds"]
    # For Sale 5 to 10 ac	Sold 5 to 10 ac
    row[34].value = counties_states[row_pos][2]["5-10"]["on-market"]
    row[35].value = counties_states[row_pos][2]["5-10"]["solds"]
    # For Sale 10 to 20 ac	Sold 10 to 20 ac
    row[45].value = counties_states[row_pos][2]["10-20"]["on-market"]
    row[46].value = counties_states[row_pos][2]["10-20"]["solds"]
    # For Sale 20 to 50 ac	Sold 20 to 50 ac
    row[56].value = counties_states[row_pos][2]["20-50"]["on-market"]
    row[57].value = counties_states[row_pos][2]["20-50"]["solds"]
    # For Sale 50 to max	Sold 50 to max
    row[67].value = counties_states[row_pos][2]["50+"]["on-market"]
    row[68].value = counties_states[row_pos][2]["50+"]["solds"]

    row_pos = row_pos + 1

wb.save("new.xlsx")
