# Atualizar Dashboard Portal Power BI

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import time
from datetime import datetime

#abrir o chrome
options = Options()
options.headless = True
driver = webdriver.Chrome('C:\\chromedriver')#Colocar o endereço do arquivo chromedriver
driver.maximize_window()

#acessar o portal
driver.get("https://portal.office.com")

#entrar com o usuário
time.sleep(10)
username = driver.find_element(By.NAME, "loginfmt")
username.send_keys("email@email.com.br")#Digitar o email
avancar = driver.find_element(By.ID, "idSIButton9")
avancar.click()

#digitar a senha
time.sleep(10)
password = driver.find_element(By.NAME, "passwd")
password.send_keys("123456")#Ddigitar a senha
avancar = driver.find_element(By.ID, "idSIButton9")
avancar.click()
avancar = driver.find_element(By.ID, "idSIButton9")
avancar.click()

#abrir uma nova guia - Disponibilidade
new_url = "https://app.powerbi.com/groups/xxx/details" #Colocar a URL detalhes do Painel
driver.execute_script("window.open('');") 
driver.switch_to.window(driver.window_handles[1])
driver.get(new_url) 

#atualizar dashboard - Disponibilidade
time.sleep(15)
update = driver.find_element(By.XPATH, '//*[@id="content"]/dataset-related-reports-container/dataset-action-bar/action-bar/action-button[2]/button')
update.click()
updatenow = driver.find_element(By.XPATH, '//*[@id="mat-menu-panel-10"]/div/button[1]')
updatenow.click()

#salvar log
time.sleep(15)
with open(r"C:\Log.txt", "a") as myfile: #colocar o local da criação do arquivo txt
    myfile.write("Atualização Realizada em "+ datetime.now().strftime("%d/%m/%Y ") + "as " + time.strftime("%H:%M:%S") + chr(13))

#Fechar Chrome
time.sleep(15)
driver.quit()
