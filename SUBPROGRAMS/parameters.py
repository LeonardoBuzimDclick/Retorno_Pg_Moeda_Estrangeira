import yagmail
import yaml
import logging
import os
import sys
import psutil
import signal
from timeit import default_timer as timer
import pdfplumber

from clicknium import locator, ui, clicknium as cc
#from clicknium import clicknium as cc
from time import sleep, time
import subprocess
import pyautogui
from datetime import datetime, timedelta
import re
import pygetwindow as gw
import pyperclip
import json
import openpyxl
import glob
import shutil

def verificar_arquivo_vazio():
    tamanho_arquivo = os.path.getsize(r'C:\Projetos_Retornos\Retorno_Pg_Moeda_Estrangeira\mes.txt')
    
    return tamanho_arquivo == 0

def ler_arquivo():
    with open(r'C:\Projetos_Retornos\Retorno_Pg_Moeda_Estrangeira\mes.txt', 'r') as arquivo:
        conteudo = arquivo.read()
    
    return conteudo


def limpar_arquivo():
    with open(r'C:\Projetos_Retornos\Retorno_Pg_Moeda_Estrangeira\mes.txt', 'w') as arquivo:
        arquivo.truncate()


pyautogui.FAILSAFE = False


base_path = os.getcwd()
log_path = f'{base_path}/LOG'
excel_path = f'{base_path}/EXCEL'
excel_erro_path = f'{base_path}/EXCEL_ERRO'
email_path = f'{base_path}/EMAIL'
image_path = f'{base_path}/IMAGES'
print_path = f'{base_path}/PRINTS'

with open('config.yaml', 'r', encoding='utf=8') as params:
    config = yaml.safe_load(params)

email_login = config['email']['login']
email_password = config['email']['password']
email_receivers = config['email']['receivers']
log_level = config['logLevel']

yag = yagmail.SMTP(user=email_login, password=email_password)
processList = ["RM.exe"]


## log configs ##
today = datetime.today().strftime('%Y-%m-%d')
log_file_name = datetime.now().strftime('%H_%M_%S')

if  os.path.exists(f'{log_path}/{today}') == False:
    os.makedirs(f'{log_path}/{today}')

logging.basicConfig(level=log_level, format='%(asctime)s; %(levelname)s; %(module)s.%(funcName)s.%(lineno)d: %(message)s',
                    handlers=[
                        logging.FileHandler(os.path.join(f'{log_path}/{today}/{log_file_name}.log'), mode='w', encoding='utf-8', delay=False),
                        logging.StreamHandler(sys.stdout)
                    ])



user = 'integracao'
password = 'arrobaint;321'
inv_palavra = 'INV'
nome_bot = 'Pagamento Moeda Estrangeira'
texto_msg = 'Sucesso na execução do Bot de Pagamento de Moeda Estrangeira.'

#localRM = 'C:\Totvs\CLOUD HML\RM.exe'
localRM = r'C:\Totvs\RM.NET\RM.exe'

#pasta_local = r"\\Riofs\dados\Documentos Financeiro\Extratos bancarios\Ebury\2024\01 - Janeiro"
pasta_original = r"\\Riofs\dados\Documentos Financeiro\Extratos bancarios\Ebury\2024"


data_atual = datetime.now()
mes_atual = data_atual.month

check_mes = verificar_arquivo_vazio()
if check_mes != True:
    mes_atual = int(ler_arquivo())


pasta_local = f'{pasta_original}\\{config['mes'][mes_atual]}'
#pasta_local = f'{pasta_original}\\03 - Março'
logging.info(f'Pasta local: {pasta_local}')


limpar_arquivo()


## EXCEL CONFIG ##
if os.path.exists(f'{excel_path}/{today}') == False:
    os.makedirs(f'{excel_path}/{today}')
writer = f'{excel_path}/{today}/{log_file_name}.xlsx'

if os.path.exists(f'{excel_erro_path}/{today}') == False:
    os.makedirs(f'{excel_erro_path}/{today}')
writer_erro = f'{excel_erro_path}/{today}/PROCESSOS COM ERRO EM - {log_file_name}.xlsx'


##  IMAGENS  ##
integracao_bancaria = f'{image_path}/integracao_bancaria.PNG'
boleto = f'{image_path}/boleto.PNG'
button_ok = f'{image_path}/button_ok.PNG'
button_ok_2 = f'{image_path}/button_ok_2.PNG'
filter_clear = f'{image_path}/filter_clean.PNG'
lancamentos = f'{image_path}/lancamentos.PNG'
forma_pag = f'{image_path}/forma_pagamento.PNG'
boleto_min = f'{image_path}/boleto_minusculo.PNG'
boleto_min_des = f'{image_path}/boleto_min_desbili.PNG'
exec_sucesso = f'{image_path}/exec_success.PNG'
boleto_min_hab = f'{image_path}/boleto_min_habili.PNG'
table_empty_img = f'{image_path}/table_empty.PNG'
concessionarias_img = f'{image_path}/concessionarias.PNG'
novo_filtro_img = f'{image_path}/novo_filtro.PNG'
licencas_excedidas = f'{image_path}/licencas_excedidas.PNG'
licencas_excedidas2 = f'{image_path}/licencas_excedidas2.PNG'



valor = f'{image_path}/valores.PNG'
baixa_success = f'{image_path}/baixa_success.PNG'

error_process_lanc = f'{image_path}/erro_lancamento.PNG'
pix = f'{image_path}/pix.PNG'
pix_transf = f'{image_path}/pix_transferencia.PNG'
tipo_chave_pix = f'{image_path}/tipo_chave_pix.PNG'
chave_pix = f'{image_path}/chave_pix.PNG'
pix_subli = f'{image_path}/pix_sublinhado.PNG'
sem_pix = f'{image_path}/sem_opcao_pix.PNG'
pix_desabilitado = f'{image_path}/tres_pontos_desabilitado.PNG'



