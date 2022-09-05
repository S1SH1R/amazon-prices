import requests
from bs4 import BeautifulSoup 
from glob import glob
from time import sleep
import pandas as pd
from datetime import datetime
import openpyxl as op

HEADERS = ({'User-Agent':
            'Chrome/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.36',
            'Accept-Language': 'en-US, en;q=0.5'})

prod_tracker = pd.read_csv('trackers.csv', sep=',')
prod_tracker_URLS = prod_tracker.Url
log = pd.DataFrame() #creates a dataframe 
date = datetime.now().strftime('%Y-%m-%d  %H:%M') #date and time
only_Date = datetime.now().strftime('%Y-%m-%d')
print(date)
print(prod_tracker_URLS[0])

#fetch the url 
page = requests.get(prod_tracker_URLS[0], headers = HEADERS)

#object contains info in the html
soup = BeautifulSoup(page.content, features = "lxml")

#get title of product
title = soup.find(id="productTitle").get_text().strip()       
print(title)

#price of the product
try:
    price = (soup.find(id="corePrice_feature_div").get_text().replace('£','').strip())
    price1 = (price[int(len(price)/2):])
except:
    price = ''
print(price1)

#product review score
review_score = soup.find(id="acrPopover").get_text().replace('out of', '/').replace('stars', '').strip()
print(review_score)

#append the csv file with the data 
logs = pd.DataFrame({'Date': date,
                    'Title': title,
                    'Price': '£' + price1,
                    'Ratings': review_score}, index=[1])

log = log.append(logs)
last_search = glob('./search_history/*.xlsx')[-1]
print(last_search)
file = pd.read_excel(last_search)
appending = file.append(log, sort=False)

appending.to_excel('search_history/SEARCH_HISTORY_{}.xlsx'.format(only_Date), index = False)

wb = op.load_workbook(last_search)
worksheet = wb.active
for col in worksheet.columns:
     max_length = 0
     column = col.column_letter # Get the column name
     for cell in col:
         try: # Necessary to avoid error on empty cells
             if len(str(cell.value)) > max_length:
                 max_length = len(str(cell.value))
         except:
             pass
     adjusted_width = (max_length + 2) * 1.2
     worksheet.column_dimensions[column].width = adjusted_width
