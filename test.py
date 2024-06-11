import requests
from bs4 import BeautifulSoup
import pandas as pd

PricseL = "https://www.spiritsandwine.lv/lv/absints/xenta-absenta-2683"
soup = BeautifulSoup(requests.get(PricseL).content, "lxml")
InfoBoks = soup.find("div", class_="page-product-container container-fluid")
Name = InfoBoks.find("h1", class_="product-title").text.strip()
print(Name)
