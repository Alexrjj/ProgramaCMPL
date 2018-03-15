import os
import openpyxl
import time
import shutil

if os.path.isfile('pastas.xlsx'):  # Verifica se o arquivo existe, caso sim, deleta a versão antiga e cria nova cópia
    os.remove('pastas.xlsx')
    shutil.copy2('pastas - template.xlsx', 'pastas.xlsx')
    os.startfile('pastas.xlsx')
else:  # Caso não existe, apenas cria nova cópia
    shutil.copy2('pastas - template.xlsx', 'pastas.xlsx')
    os.startfile('pastas.xlsx')

input('Enter any key to continue...')

wb = openpyxl.load_workbook('pastas.xlsx')
for sheet in wb.worksheets:
    path = sheet['A2'].value
    for row in sheet.iter_cols(min_row=2, min_col=2, max_col=2):
        print('Change the date in xlsm file to ' + sheet['A2'].value)
        os.startfile('FluxoObraPadrao4.0.xlsm')
        for x in range(0, 1000):
            try:
                xls = open('FluxoObraPadrao4.0.xlsm', 'r+')
                if xls:
                    xls.close()
                    break
            except IOError:
                time.sleep(1)
                continue
        for cell in row:
            print('Creating folder... ' + str(cell.value))
            os.makedirs(os.path.join(path, str(cell.value)))
            shutil.copy2('FluxoObraPadrao4.0.xlsm', os.path.join(path, str(cell.value)))
            os.rename(os.path.join(path, str(cell.value), 'FluxoObraPadrao4.0.xlsm'), (os.path.join(path, str(cell.value), str(cell.value) + '.xlsm')))
            time.sleep(3)
input('Fim da execução.')