from selenium import webdriver
from chromedriver_py import binary_path
import time
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import pyautogui


chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory" : "C:\\Users\\jvict\\Documents\\seven\\downloads_malha\\diario_oficial2023"}
chromeOptions.add_experimental_option("prefs",prefs)
svc = webdriver.ChromeService(executable_path=binary_path)
driver = webdriver.Chrome(service=svc, options=chromeOptions)

driver.get('https://diario.imprensaoficial.al.gov.br/edicoes')
time.sleep(2)
botao_2020s = driver.find_element(By.XPATH, '//*[@id="6"]')
botao_2020s.click()
time.sleep(2)
#fazendo 2020
ano = driver.find_element(By.XPATH, '//*[@id="2023"]')
ano.click()
time.sleep(2)
for x in range(12):
    mes = driver.find_element(By.XPATH, '/html/body/div/div/div[2]/div/div[2]/div/div/div[3]/div/div['+str(x+1)+']/button')
    mes.click()
    time.sleep(2)
    rows = driver.find_elements(
    By.XPATH, '/html/body/div/div/div[2]/div/div[3]/div/table/tbody/tr')
    rowsCount = len(rows)
    print(rowsCount)
    time.sleep(2)
    for i in range(rowsCount):
        baixador = driver.find_element(By.XPATH, '/html/body/div/div/div[2]/div/div[3]/div/table/tbody/tr['+str(i+1)+']/td[4]/a[2]')
        baixador.click()
        time.sleep(3)