from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import time
import csv
import datetime
from urllib.parse import urlparse
import pandas as pd

current_time = datetime.datetime.now()

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

# URL to scrape
url = 'https://www.tugo.co/muebles/muebles-auxiliares/auxiliares-de-bano'

# Open the URL using Selenium
driver.get(url)

# Parse the URL
parsed_url = urlparse(url)

# Get the path component of the URL
path = parsed_url.path

# Split the path by '/'
path_parts = path.split('/')

# The last part of the path should be 'auxiliares-de-bano'
product_style = path_parts[-1]

# Allow time for the JavaScript to execute and content to load
time.sleep(30)  # Adjust the sleep time as needed

# Get the page source and parse it with BeautifulSoup
html_content = driver.page_source
soup = BeautifulSoup(html_content, 'html.parser')

# Extract the title of the page
title = soup.find('title').text
print('Title:', title)

# Extract all product information (example extraction)
products = []
products.append(['Product_Url','Product Name', 'Old_Price', 'Current_Price', 'Prodcut_style'])
# Assuming products are listed in divs with a specific class
product_divs = soup.find_all('div', class_='vtex-search-result-3-x-galleryItem')
for div in product_divs:
    product_url = div.find('a').get('href') if div.find('a').get('href') else 'N/A'
    product_name = div.find('h3').text if div.find('h3') else 'N/A'
    product_old_price = div.find('div', class_='vtex-store-components-3-x-listPrice').text if div.find('div', class_='vtex-store-components-3-x-listPrice') else 'N/A'
    product_cur_price = div.find('div', class_='vtex-store-components-3-x-sellingPrice').text if div.find('div', class_='vtex-store-components-3-x-sellingPrice') else 'N/A'
    products.append([product_url, product_name, product_old_price, product_cur_price, product_style])

# Close the WebDriver
driver.quit()

# Define the CSV file name
csv_file = current_time.strftime("%Y-%m-%d_%H-%M-%S.csv")
exl_file = current_time.strftime("%Y-%m-%d_%H-%M-%S.xlsx")

# Save the data to a CSV file
with open(csv_file, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)
    # Write the header
    # writer.writerow(['Product_Url','Product Name', 'Old_Price', 'Current_Price', 'Prodcut_style'])
    # Write the product data
    writer.writerows(products)

df = pd.DataFrame(products)   
with pd.ExcelWriter(exl_file, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    
    # Access the workbook and worksheet
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Example: Set the column width and format
    format1 = workbook.add_format({'num_format': '0.00'})  # Format for numbers
    worksheet.set_column('B:B', 12, format1)  # Set width of column B
    
    # Example: Add a header format
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7E4BC',
        'border': 1
    })
    
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)


print(f'Data has been written to {csv_file}')
print(f'Data has been written to {exl_file}')
