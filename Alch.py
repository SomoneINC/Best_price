import pandas as pd
import random
import requests
from bs4 import BeautifulSoup
import os

class PC():
    SpriritListSections = []
    SpiritDrinkList = []

link = "https://www.spiritsandwine.lv/lv/"
req = requests.get(link)
x = 0
soup = BeautifulSoup(req.content, 'html.parser')
selection = soup.find_all('div', class_="col-sm-5 col-md-4 col-xl-3 mb-1")
for items in selection:
    find_a = items.find('a')
    if find_a:
        Link = find_a['href'][4:]
        PC.SpriritListSections.append(link + Link)

req = requests.get(PC.SpriritListSections[0])
soup = BeautifulSoup(req.content, 'html.parser')


while True:
    drink_box = soup.find_all('div', class_="product-container")
    for items in drink_box:
        find_a = items.find('a')
        if find_a:
            Link = find_a['href']
            PC.SpiritDrinkList.append(link +Link)
    next_button = soup.find('a', class_="btn-next")
    if next_button:
        x += 1
        print(x)
        req = requests.get(link + next_button['href'])
    elif not next_button:
        break

for link in PC.SpiritDrinkList:
    print(link)
"""
# Create a DataFrame with a link
data = {'Yarn': ["Name"]}
df = pd.DataFrame(data)
random_number = random.randint(100000, 999999)
# Create a Pandas Excel writer using XlsxWriter as the engine
file_name = f"{random_number}.xlsx"
writer = pd.ExcelWriter(file_name, engine='xlsxwriter')

# Write the DataFrame to the Excel file
df.to_excel(writer, sheet_name='Sheet1', index=False)

# Get the xlsxwriter workbook and worksheet objects
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Add a hyperlink format to the workbook
url_format = workbook.add_format({'color': 'blue', 'underline': True})

# Add the hyperlink to the cell
worksheet.write_url('B1', link, url_format, string='yes')

# Save the Excel file
writer._save()
os.system(f'start excel {file_name}')"""
