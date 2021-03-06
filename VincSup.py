from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, NoAlertPresentException
from selenium import webdriver
import openpyxl
from selenium.webdriver.support.ui import Select
import time
import os
import shutil


if os.path.isfile('VincSupVist.txt'):  # Verifica se o arquivo existe, caso sim, deleta a versão antiga e cria nova cópia
    os.remove('VincSupVist.txt')
    open('VincSupVist.txt', 'a').close()
else:  # Caso não existe, apenas cria nova cópia
    open('VincSupVist.txt', 'a').close()

os.startfile('VincSupVist.txt')

input('Enter any key to continue...')

#  Acessa os dados de login fora do script, salvo numa planilha existente, para proteger as informações de credenciais
dados = openpyxl.load_workbook('C:\\gomnet.xlsx')
login = dados['Plan1']
url = 'http://gomnet.ampla.com/'
urlVincSup = 'http://gomnet.ampla.com/vistoria/vincularSupervisor.aspx'
consulta = 'http://gomnet.ampla.com/ConsultaObra.aspx'
username = login['A1'].value
password = login['A2'].value

# --------------- Headless Mode -------------------------
# chromeOptions = webdriver.ChromeOptions()
# prefs = {"download.default_directory" : os.getcwd(),
#          "download.prompt_for_download": False}
# chromeOptions.add_experimental_option("prefs",prefs)
# chromeOptions.add_argument('--headless')
# chromeOptions.add_argument('--window-size= 1600x900')
# driver = webdriver.Chrome(chrome_options=chromeOptions)
# -------------------------------------------------------

driver = webdriver.Chrome()

if __name__ == '__main__':
    driver.get(url)
    # Insere login/senha e entra no sistema
    uname = driver.find_element_by_name('txtBoxLogin')
    uname.send_keys(username)
    passw = driver.find_element_by_name('txtBoxSenha')
    passw.send_keys(password)
    submit_button = driver.find_element_by_id('ImageButton_Login').click()

    # Acessa a página de vinculação de supervisor
    driver.get(urlVincSup)

    # Insere o número da Sob em seu respectivo campo e realiza a busca
    with open('VincSupVist.txt') as data:
        datalines = (line.strip('\r\n') for line in data)
        for line in datalines:
            driver.find_element_by_id('ctl00_ContentPlaceHolder1_TextBox_SOB').clear()
            sob = driver.find_element_by_id('ctl00_ContentPlaceHolder1_TextBox_SOB')
            sob.send_keys(line)
            # Pesquisa pela sob 03 vezes
            driver.find_element_by_id('ctl00_ContentPlaceHolder1_ImageButton_Pesquisar').click()
            try:
                #
                supervisor = Select(driver.find_element_by_id('ctl00_ContentPlaceHolder1_gridViewTarefas_ctl02_ddlSupervisor'))
                supervisor.select_by_visible_text('MESSIAS JOSE DE FARIA')
                driver.find_element_by_id('ctl00_ContentPlaceHolder1_Button_VincularSupervisor').click()
                time.sleep(1)
                try:
                    alert = driver.switch_to_alert()
                    alert.accept()
                    webdriver.ActionChains(driver).send_keys(Keys.SPACE).perform()
                except NoAlertPresentException:
                    continue
            except NoSuchElementException:
                log = open("VincSupOutput.txt", "a")
                log.write(line + ' não vinculada. ' + '\n')
                log.close()
                continue