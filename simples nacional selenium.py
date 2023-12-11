import undetected_chromedriver as uc
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
 
options = uc.ChromeOptions()
options.headless = False
prefs = {"download.default_directory" : "C:\\Users\\jvict\\Documents\\seven\\downloads_malha"}
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_experimental_option("prefs",prefs)

driver = uc.Chrome(options=options)
#driver.delete_all_cookies()
driver.maximize_window()


driver.get('https://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATSPO/dasnsimei.app')

cnpj = '45135775000102'

barra_de_pesquisa = driver.find_element(By.XPATH, '/html/body/div[1]/main/div/div/div[2]/div/div/div/form/div[1]/div/div[1]/input')
barra_de_pesquisa.send_keys(cnpj)
time.sleep(2)


continuar = driver.find_element(By.XPATH,'/html/body/div[1]/main/div/div/div[2]/div/div/div/form/div[1]/div/div[3]/button')
continuar.click()

time.sleep(30)
