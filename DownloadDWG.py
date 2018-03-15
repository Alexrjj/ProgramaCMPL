import os
import time
from selenium.common.exceptions import NoSuchElementException
from selenium import webdriver
import openpyxl

#  Acessa os dados de login fora do script, salvo numa planilha existente, para proteger as informações de credenciais
dados = openpyxl.load_workbook('C:\\gomnet.xlsx')
login = dados['Plan1']
url = 'http://gomnet.ampla.com/'
url_dwnld = 'http://gomnet.ampla.com/Upload.aspx?numsob='
username = login['A1'].value
password = login['A2'].value

# Configurações do Browser
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
    with open('sobs.txt') as data:
        datalines = (line.rstrip('\r\n') for line in data)
        for line in datalines:
            driver.get(url_dwnld + line)
            window_before = driver.window_handles[0]
            try:
                # Busca pela versão mais atual do dwg da sob
                try:
                    rev_09 = driver.find_element_by_partial_link_text('REV.09.dwg')
                    if rev_09.is_displayed():
                        rev_09.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev_08 = driver.find_element_by_partial_link_text('REV.08.dwg')
                    if rev_08.is_displayed():
                        rev_08.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev_07 = driver.find_element_by_partial_link_text('REV.07.dwg')
                    if rev_07.is_displayed():
                        rev_07.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev_06 = driver.find_element_by_partial_link_text('REV.06.dwg')
                    if rev_06.is_displayed():
                        rev_06.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev_05 = driver.find_element_by_partial_link_text('REV.05.dwg')
                    if rev_05.is_displayed():
                        rev_05.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev_04 = driver.find_element_by_partial_link_text('REV.04.dwg')
                    if rev_04.is_displayed():
                        rev_04.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev_03 = driver.find_element_by_partial_link_text('REV.03.dwg')
                    if rev_03.is_displayed():
                        rev_03.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev_02 = driver.find_element_by_partial_link_text('REV.02.dwg')
                    if rev_02.is_displayed():
                        rev_02.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev_01 = driver.find_element_by_partial_link_text('REV.01.dwg')
                    if rev_01.is_displayed():
                        rev_01.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra9 = driver.find_element_by_partial_link_text('REVOBRA_09.dwg')
                    if revObra9.is_displayed():
                        revObra9.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra8 = driver.find_element_by_partial_link_text('REVOBRA_08.dwg')
                    if revObra8.is_displayed():
                        revObra8.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra7 = driver.find_element_by_partial_link_text('REVOBRA_07.dwg')
                    if revObra7.is_displayed():
                        revObra7.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra6 = driver.find_element_by_partial_link_text('REVOBRA_06.dwg')
                    if revObra6.is_displayed():
                        revObra6.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra5 = driver.find_element_by_partial_link_text('REVOBRA_05.dwg')
                    if revObra5.is_displayed():
                        revObra5.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra4 = driver.find_element_by_partial_link_text('REVOBRA_04.dwg')
                    if revObra4.is_displayed():
                        revObra4.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra3 = driver.find_element_by_partial_link_text('REVOBRA_03.dwg')
                    if revObra3.is_displayed():
                        revObra3.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra2 = driver.find_element_by_partial_link_text('REVOBRA_02.dwg')
                    if revObra2.is_displayed():
                        revObra2.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra1 = driver.find_element_by_partial_link_text('REVOBRA_01.dwg')
                    if revObra1.is_displayed():
                        revObra1.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    revObra = driver.find_element_by_partial_link_text('REVOBRA.dwg')
                    if revObra.is_displayed():
                        revObra.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev9 = driver.find_element_by_partial_link_text('REV09.dwg')
                    if rev9.is_displayed():
                        rev9.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev8 = driver.find_element_by_partial_link_text('REV08.dwg')
                    if rev8.is_displayed():
                        rev8.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev7 = driver.find_element_by_partial_link_text('REV07.dwg')
                    if rev7.is_displayed():
                        rev7.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev6 = driver.find_element_by_partial_link_text('REV06.dwg')
                    if rev6.is_displayed():
                        rev6.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev5 = driver.find_element_by_partial_link_text('REV05.dwg')
                    if rev5.is_displayed():
                        rev5.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev4 = driver.find_element_by_partial_link_text('REV04.dwg')
                    if rev4.is_displayed():
                        rev4.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev3 = driver.find_element_by_partial_link_text('REV03.dwg')
                    if rev3.is_displayed():
                        rev3.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev2 = driver.find_element_by_partial_link_text('REV02.dwg')
                    if rev2.is_displayed():
                        rev2.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    rev1 = driver.find_element_by_partial_link_text('REV01.dwg')
                    if rev1.is_displayed():
                        rev1.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
                try:
                    dwg = driver.find_element_by_partial_link_text('dwg')
                    if dwg.is_displayed():
                        dwg.click()
                        fecha_janela()
                        continue
                except NoSuchElementException:
                    pass
            except NoSuchElementException:
                log = open("ErroSobs.txt", "a")
                log.write(line + "\n")
                log.close()
                continue
