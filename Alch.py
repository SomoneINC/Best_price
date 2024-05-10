from openpyxl import Workbook
from openpyxl import load_workbook
from selenium import webdriver
import threading
import random
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

    SpritSelection = []
    service = Service(ChromeDriverManager().install())
    link = "https://www.spiritsandwine.lv/"
    response = requests.get(link)
    soup = BeautifulSoup(response.content, "html.parser")
    a_tag = soup.find_all("a", class_="dropdown-item")
    for a in a_tag:
        if a["href"][:4] != "http": 
            print(a["href"])
   

    """
    workbook = Workbook()
    sheet = workbook.create_sheet(title="Locations")
    for section in SpritSelection :
        clean_section = section.translate(str.maketrans("", "", string.punctuation))
        sheet = workbook.create_sheet(title=clean_section)
    workbook.save("Locations.xlsx")"""
    print("Done")
    time.sleep(2000)
    
   

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
