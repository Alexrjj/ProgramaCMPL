import os
import time
from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
import openpyxl
import datetime

#  Acessa os dados de login fora do script, salvo numa planilha existente, para proteger as informações de credenciais
dados = openpyxl.load_workbook('C:\\gomnet.xlsx')
login = dados['Plan1']
url = 'http://gomnet.ampla.com/'
url_dwnld = 'http://gomnet.ampla.com/Upload.aspx?numsob='
username = login['A1'].value
password = login['A2'].value

# Configurações do Chrome
chromeOptions = webdriver.ChromeOptions()
prefs = {"download.default_directory": os.getcwd(),
         "download.prompt_for_download": False}
chromeOptions.add_experimental_option("prefs", prefs)
# chromeOptions.add_argument('--headless')
driver = webdriver.Chrome(chrome_options=chromeOptions)

if __name__ == '__main__':
    driver.get(url)
    # Faz login no sistema
    uname = driver.find_element_by_name('txtBoxLogin')
    uname.send_keys(username)
    passw = driver.find_element_by_name('txtBoxSenha')
    passw.send_keys(password)
    submit_button = driver.find_element_by_id('ImageButton_Login').click()

    def fecha_janela():
        window_after = driver.window_handles[1]
        driver.switch_to_window(window_after)
        driver.close()
        driver.switch_to_window(window_before)
        time.sleep(5)

    # Insere o número da Sob em seu respectivo campo e realiza a busca
    mes = datetime.datetime.now().strftime("%m")
    with open('DownloadSGD.txt') as data:
        datalines = (line.rstrip('\r\n') for line in data)
        for line in datalines:
            driver.get(url_dwnld + line)
            window_before = driver.window_handles[0]
            try:
                try:
                    pdfs = len(driver.find_elements_by_partial_link_text('-' + mes + '-18'))
                    for pdf in range(pdfs):
                            driver.find_elements_by_partial_link_text('-' + mes + '-18')[pdf].click()
                            fecha_janela()
                            continue
                except NoSuchElementException:
                    pass
            except NoSuchElementException:
                log = open("DownloadSGDErro.txt", "a")
                log.write(line + "\n")
                log.close()
                continue
