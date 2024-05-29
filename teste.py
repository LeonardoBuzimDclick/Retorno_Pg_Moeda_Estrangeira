import pygetwindow as gw
import pyautogui
import time
import re
from clicknium import clicknium as cc
from SUBPROGRAMS.parameters import *
from SUBPROGRAMS.functions import *
from clicknium import locator, ui
from clicknium import clicknium as cc
import time
import subprocess
import pyautogui
from datetime import datetime, timedelta
import re
import pygetwindow as gw
import pyperclip



# wait for image with pyautogui
'''def waitForImage(image, timeout, name):
    inicialTime = time.time()

    while True:
        # try to find image on screen
        location = pyautogui.locateOnScreen(image=image, confidence=0.9)

        if location is not None:
            # Image found
            logging.info(f'{name} image found.')
            return location

        # checking if timeout is over
        currentTime = time.time()
        if currentTime - inicialTime >= timeout:
            # timeout is over and the image was not found.
            logging.info(f'{name} image not found.')
            return None'''



def emailSuccess(nome_bot, texto_msg):
    with open(f'{email_path}/emailSuccess.html', 'r', encoding='utf=8') as fileSuccess:
        templateSuccess = fileSuccess.read()
        htmlSuccess = templateSuccess.format(automacao = nome_bot, mensagem = texto_msg)
    yag.send(to=str(email_receivers).split(','), subject=f"SUCESSO - BOT REMESSAS PIX",
    contents=htmlSuccess)
    #attachments=f'{log_path}/{today}/{log_file_name}.log')
    logging.info('Success e-mail sent.')



'''close = cc.wait_appear(locator.rm.button_btncancel, wait_timeout=60)
close.click()
logging.debug(f'FECHAR')'''







'''select_valores = waitForImage(image = valor, timeout=30, name = 'INTEGRAÇÃO BANCÁRIA')
if select_valores is not None:
    sleep(2)
    pyautogui.click(select_valores)
    logging.info('CLICK EM VALOR')

sleep(1)

for i in range (4):
    sleep(0.1)
    pyautogui.press('tab')
logging.info('4 TAB')'''


'''select_filter = cc.wait_appear(locator.rm.select_filter, wait_timeout=60)
select_filter.click()
sleep(2)

filter_global = cc.wait_appear(locator.rm.group_meus_filtros, wait_timeout=60)
filter_global.click()
sleep(5)

#iltro = 'histórico'
#filtro = 'Valor Original'
filtro = 'Número do documento'

filter_global.set_text(f'{filtro}')
logging.info(f'ESCREVE {filtro} NO FILTRO')
sleep(1)

execute = cc.wait_appear(locator.rm.button_buttonexec, wait_timeout=15)
execute.click()
logging.info('CLICK - BOTÃO DE SELECAO DE FILTRO')'''

moeda = 'GBP'

padrao_moeda = ['NOK','EUR','USD','GBP','BRL']
if moeda in padrao_moeda:
    print('1')
else:
    print('2')