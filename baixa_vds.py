import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

url = "https://oss.telebras.com.br/cpqdom-web/login.xhtml"
url_tabela = "https://oss.telebras.com.br/cpqdom-web/operation/OrderQueryList.xhtml"
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-extensions")
options.add_argument("--disable-gpu")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--no-sandbox")
options.add_argument("--force-device-scale-factor=0.75")
servico = Service(ChromeDriverManager().install())

prefs = {
    "download.default_directory": os.getcwd(),
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)

chrome = webdriver.Chrome(service=servico, options=options)
chrome.get(url)

wait = WebDriverWait(chrome, 300)
wait.until(EC.url_changes(url))

chrome.get(url_tabela)

chrome.find_element(By.XPATH, '/html/body/form[2]/div/div/div/div/div[1]/div/button[2]').click()
time.sleep(1)

while len(chrome.find_elements(By.ID, 'dataTableFormId:DataTableId:j_idt252')) < 1:
    time.sleep(1)
chrome.find_element(By.ID, 'dataTableFormId:DataTableId:j_idt252').click()

time.sleep(5)
elemento_filtro = chrome.find_element(By.XPATH, '/html/body/form[2]/div/div/div/div/div[2]/div/table/thead/tr/th[1]/input')
elemento_filtro.send_keys('VDS')
chrome.find_element(By.XPATH, '/html/body/form[2]/div/div/div/div/div[2]/div/table/thead/tr/th[8]/span[2]').click()
time.sleep(5)
chrome.find_element(By.XPATH, '/html/body/form[2]/div/div/div/div/div[1]/a[1]').click()
time.sleep(10)