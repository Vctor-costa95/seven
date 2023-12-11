import undetected_chromedriver as uc
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
#PQ Q NAO TAVA FUNCIONANDO? PQ TEM QUE BAIXAR O NOVO CHROMEDRIVER NO SITE https://googlechromelabs.github.io/chrome-for-testing/
#quando baixar ele colocar no service e cabou.

service = webdriver.ChromeService(executable_path='C:/Users/jvict/Downloads/chromedriver-win64/chromedriver-win64/chromedriver.exe')
options = webdriver.ChromeOptions()
options.headless = False
prefs = {"download.default_directory" : "C:\\Users\\jvict\\Documents\\seven\\downloads_malha"}
options.add_experimental_option("prefs",prefs)

driver = webdriver.Chrome(options=options)
driver.delete_all_cookies()
driver.maximize_window()
actions = ActionChains(driver)
driver.get('https://www.webmotors.com.br/carros-usados/al-maceio?estadocidade=Alagoas%20-%20Macei%C3%B3&tipoveiculo=carros-usados&localizacao=-9.6576274,-35.7244726x0km')
a = True
while a == True:
    try:
        final = driver.find_element(By.XPATH,'/html/body/div[1]/main/div[1]/div[3]/div[2]/div/div[1]/div/div[130]')
        a = False
    except:
        actions.send_keys("end")
        time.sleep(5)
time.sleep(5)

#/html/body/div[1]/main/div[1]/div[3]/div[2]/div/div[1]/div/div[1]/div/div[2]/a[2]/strong
#/html/body/div[1]/main/div[1]/div[3]/div[2]/div/div[1]/div/div[2]/div/div[2]/a[2]/strong
#/html/body/div[1]/main/div[1]/div[3]/div[2]/div/div[1]/div/div[3]/div/div[2]/a[2]/strong
#/html/body/div[1]/main/div[1]/div[3]/div[2]/div/div[1]/div/div[130]/div/div[2]/a[2]/strong

#/html/body/div[1]/main/div[1]/div[3]/div[2]/div/div[1]/div/div[130]
#/html/body/div[1]/main/div[1]/div[3]/div[2]/div/div[1]/div/div[81]