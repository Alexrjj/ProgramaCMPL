﻿import os
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import selenium
from selenium import webdriver
import openpyxl

#  Acessa os dados de login fora do script, salvo numa planilha existente, para proteger as informações de credenciais
dados = openpyxl.load_workbook('C:\\gomnet.xlsx')
login = dados['Plan1']
url = 'http://gomnet.ampla.com/'
consulta = 'http://gomnet.ampla.com/ConsultaObra.aspx'
username = login['A1'].value
password = login['A2'].value

chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory" : os.getcwd(),
         "download.prompt_for_download": False}
chromeOptions.add_experimental_option("prefs",prefs)
chromeOptions.add_argument('--headless')
chromeOptions.add_argument('--window-size= 1600x900')
driver = webdriver.Chrome(chrome_options=chromeOptions)

# driver = webdriver.Chrome()
if __name__ == '__main__':
    driver.get(url)
    # Faz login no sistema
    uname = driver.find_element_by_name('txtBoxLogin')
    uname.send_keys(username)
    passw = driver.find_element_by_name('txtBoxSenha')
    passw.send_keys(password)
    submit_button = driver.find_element_by_id('ImageButton_Login').click()

    # Acessa a página de Consulta de Obras
    driver.get('http://gomnet.ampla.com/ConsultaObra.aspx')

    # Insere o número da Sob em seu respectivo campo e realiza a busca
    sob = driver.find_element_by_id('ctl00_ContentPlaceHolder1_TextBox_NumSOB')

    with open('ExtraiSobStatus.txt') as data:
        datalines = (line.rstrip('\r\n') for line in data)
        for line in datalines:
            driver.find_element_by_id('ctl00_ContentPlaceHolder1_TextBox_NumSOB').clear()
            sob = driver.find_element_by_id('ctl00_ContentPlaceHolder1_TextBox_NumSOB')
            sob.send_keys(line)
            driver.find_element_by_id('ctl00_ContentPlaceHolder1_ImageButton_Enviar').click()
            try:
                # Verifica se a sob está no status desejado
                numSob = driver.find_element_by_xpath(
                    '/html/body/form/table/tbody/tr[4]/td/div[3]/table/tbody/tr[2]/td[8][contains(text(), "' + line + '")]')
                if numSob.is_displayed():
                    numSobArquivo = driver.find_element_by_xpath(
                        '//*[@id="ctl00_ContentPlaceHolder1_Gridview_GomNet1"]/tbody/tr[2]/td[8]').text
                    numStatusArquivo = driver.find_element_by_xpath(
                        '//*[@id="ctl00_ContentPlaceHolder1_Gridview_GomNet1"]/tbody/tr[2]/td[3]').text
                    log = open("ExtraiSobStatusOutput.txt", "a")
                    log.write(numSobArquivo + " " + numStatusArquivo + "\n")
                    log.close()
            except NoSuchElementException:
                continue
    print("Fim da execução.")
