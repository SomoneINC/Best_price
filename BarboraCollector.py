from openpyxl import Workbook
from selenium import webdriver
import threading
import re
import time
from bs4 import BeautifulSoup
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import requests
import threading



class A:
    Workbook = Workbook()

link = "https://www.barbora.lv/dzerieni"
response = requests.get(link)
soup = BeautifulSoup(response.content, "html.parser")

DrinkList = soup.find_all("a", class_="category-item--title")
print(soup)

if DrinkList is not None:

    print(DrinkList)
else:
    print("No element found with the specified XPath")

time.sleep(5)