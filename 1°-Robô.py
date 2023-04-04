from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import pyautogui

driver = webdriver.Chrome()

#Entrando Na pagina inicial
#driver.get("http://comexstat.mdic.gov.br/pt/home")

#Entrando Na página de municipios
driver.get("http://comexstat.mdic.gov.br/pt/municipio")


#Entrando dentro das Exportações
#ExportaçõesMunicipios = driver.find_element(By.XPATH, '/html/body/app-root/div/app-homepage-cards/div[2]/div[2]/div[2]/a/div/span')

# Clicando no elemento
#ExportaçõesMunicipios.click()

#Importação
#imp = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form"]/div[1]/div[2]/div/p-radiobutton[2]/div/div[2]')))
#imp.click()

# Localize o elemento de seleção
SelectAno = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="selectYearStart"]/div/div[1]')))
SelectAno.click()

# Aguarde até que as opções sejam exibidas
opcoes_Ano = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//*[@id="selectYearStart"]/div/div[2]/div/div[4]')))

# Percorra as opções até encontrar o valor desejado e clique nele
for opcao in opcoes_Ano:
    if opcao.text == '2000':
        opcao.click()

#Marca o checkbox De detalhamento de mês
Detalha_Mês = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form"]/div[2]/div[1]/div[4]/div/p-checkbox/div/div[2]')))
Detalha_Mês.click()

#Marcando filtros
filtro = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="form"]/div[4]/div/ng-selectize/div/div[1]/input')))
filtro.click()

#Escolhendo os Filtros
primeiro_Filtro = pyautogui.write('Cap') 
pyautogui.press('enter')
segundo_Filtro = pyautogui.write('UF')
pyautogui.press('enter')
terceiro_Filtro = pyautogui.write('Muni')
pyautogui.press('enter')

Sair_Dos_Filtros = pyautogui.press('esc')

#Selecionando os dados do primeiro filtro
First_Filter = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="form"]/div[5]/div[1]/app-input-list/div/div/ng-selectize/div/div[1]/input')))
First_Filter.click()

Primeiro = pyautogui.write('08')
pyautogui.press('enter')

Sair_Do_First_Filter = pyautogui.press('esc')
time.sleep(1.5)
Tab1 = pyautogui.press('tab')

#Segundo Filtro
Second_Filter = pyautogui.write('Ce')
pyautogui.press('enter')


#Marca o check box de Quilograma_Liquido
Quilograma_Liquido = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="form"]/div[7]/div/p-checkbox[2]/div/div[2]')))
Quilograma_Liquido.click()

#Clique em pesquisar
Localizar = WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="form"]/div[10]/button[2]')))
Localizar.click()


#Dados na vertical ou horizontal?
Vertical = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/div/app-city-form/div/div[1]/app-table-result/div[2]/div[1]/form/div/div/p-radiobutton[2]/div/div[2]')))
Vertical.click()

time.sleep(0.5)

#Baixar o Excel
Excel = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/div/app-city-form/div/div[1]/app-table-result/div[2]/div[2]/button[2]')))
Excel.click()

# Tempo que a página fica no ar após terminar todos os scripts
time.sleep(65)

# Encerrando o driver
driver.quit()


