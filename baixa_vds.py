import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

servico = Service(ChromeDriverManager().install())

url = "https://oss.telebras.com.br/cpqdom-web/login.xhtml"
url_tabela = "https://oss.telebras.com.br/cpqdom-web/operation/OrderQueryList.xhtml"
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-extensions")
options.add_argument("--disable-gpu")
# options.add_argument("--disable-dev-shm-usage")
# options.add_argument("--no-sandbox")
options.add_argument("--force-device-scale-factor=0.75")

prefs = {
    "download.default_directory": os.getcwd(),
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)

chrome = webdriver.Chrome(options=options)
chrome.get(url)

wait = WebDriverWait(chrome, 300)
wait.until(EC.url_changes(url))

chrome.get(url_tabela)

chrome.find_element(By.XPATH, '/html/body/form[2]/div/div/div/div/div[1]/div/button[2]').click()
time.sleep(1)

while len(chrome.find_elements(By.ID, 'dataTableFormId:DataTableId:j_idt466')) < 1:
    time.sleep(1)
chrome.find_element(By.ID, 'dataTableFormId:DataTableId:j_idt466').click()

time.sleep(5)
elemento_filtro = chrome.find_element(By.XPATH, '//*[@id="dataTableFormId:DataTableId:j_idt304:filter"]')
elemento_filtro.send_keys('VDS')
chrome.find_element(By.XPATH, '//*[@id="dataTableFormId:DataTableId:j_idt320"]/span[2]').click()
time.sleep(5)
chrome.find_element(By.XPATH, '//*[@id="dataTableFormId:DataTableId_paginator_top"]/a[1]/img').click()
time.sleep(10)

