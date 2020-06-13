from selenium import webdriver
from bs4 import BeautifulSoup
import pandas as pd
from selenium.webdriver.chrome.options import Options
import os

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

po = input("enter po: ")


options = Options()
options.headless = True
options.add_argument("--window-size=1920,1200")


DRIVER_PATH = r"./chromedriver_win32/chromedriver.exe"
driver = webdriver.Chrome(options=options, executable_path=resource_path(DRIVER_PATH))
driver.get("http://ioms/MFG/ModuleStatus?PO=" + po + "#!/")
source = driver.page_source
soup=BeautifulSoup(driver.page_source, features="html.parser")
percents = soup.findAll('text',class_='percentage')
modules = ["Build - ", "Test - ", "Prep - ", "Total - "]
print("Chamber Progress:\n")
for i in range(4):
    print(modules[i] + percents[i].text + "\n")
