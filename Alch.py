from openpyxl import Workbook
from openpyxl import load_workbook
from selenium import webdriver
import threading
import random
import re
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import keyboard
import string
import requests
import time
import os
import threading
import pyautogui
import pyperclip
import platform


def SpiritAndWine():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--disable-features=VizDisplayCompositor")
    chrome_options.add_argument("--disable-gpu")

    SpritSelectionLinks = []
    SpiritSelectionNames = []
    ProductNames = []
    ProductLinks = []
    ProductPrices = []
    service = Service(ChromeDriverManager().install())
    link = "https://www.spiritsandwine.lv"
    response = requests.get(link)
    x = 0	
    #Get all types of drink links
    soup = BeautifulSoup(response.content, "html.parser")
    a_tag = soup.find_all("a", class_="dropdown-item")
    for a in a_tag:
        if a["href"][:4] != "http" and x < 39: 
            SpritSelectionLinks.append(link + a["href"])
            print(link +a["href"])
            x += 1

    workbook = Workbook()

    #Get all drink names, prices and valuemes of all the drinks
    for section in SpritSelectionLinks :
        soup = BeautifulSoup(requests.get(section).content, "html.parser")
        productsCount = 1
        title = soup.find("h1", class_="category-header-info").text
        clean_title = re.sub(r"[/]", "", title)
        sheet = workbook.create_sheet(title=clean_title)
        while True:
            slecetions_box = soup.find("div", class_="row row-cols-1 row-cols-md-2 row-cols-lg-3 row-cols-xxl-4 g-0")
            Links = slecetions_box.find_all("a")
            for links in Links:
                PLink = link + links["href"]
                soup = BeautifulSoup(requests.get(PLink).content, "html.parser")
                Name = soup.find("h1", class_="product-title").text
                Sale = soup.find("div", class_="badge-sale")
                
                sheet.cell(row=productsCount, column=1).value = Name
                if Sale is not None :
                    sheet.cell(row=productsCount, column=3).value = "Akcija"
                    Price = soup.find("div", class_="actual")
                    
                    if Price is not None :
                        sheet.cell(row=productsCount, column=2).value = Price.text
                    else :
                        sheet.cell(row=productsCount, column=2).value = "Not found"
                else :
                    sheet.cell(row=productsCount, column=3).value = "Nav"
                    Price = soup.find("div", class_="product-price mr-1 text-right").text
                    sheet.cell(row=productsCount, column=2).value = Price

                sheet.cell(row=productsCount, column=4).value = PLink
                productsCount += 1 
            NextPage = soup.find("a", class_="btn-next")
            if  NextPage is not None and 'href' in NextPage.attrs :
                soup = BeautifulSoup(requests.get(link + NextPage["href"]).content, "html.parser")
                print("page next")
            else :
                break
        
        """
        boxes = slecetions_box.find_all("h2")
        for box in boxes:
            print(box.text)"""
        print(title + " done")
        SpiritSelectionNames.append(title)
    workbook.save("SpiritAndVine.xlsx")

    """
    
    
    for section in SpritSelection :
        clean_section = section.translate(str.maketrans("", "", string.punctuation))
        sheet = workbook.create_sheet(title=clean_section)
    """
    print("Done")
    
   

SpiritAndWine()


"""
# Create a new workbook
workbook = Workbook()

sheet = workbook.active
sheet.cell(row=1, column=1).value = "Hello, World!"

max_length = 0
for row in sheet.iter_rows(min_row=1, min_col=1, max_col=1):
    for cell in row:
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))

column = sheet.column_dimensions[f'{chr(0 + 65)}']
column.width = max_length

workbook.save('output.xlsx')
os.system(f'start excel {"output.xlsx"}')"""
