from SUBPROGRAMS.functions import *
from clicknium import locator, ui
from clicknium import clicknium as cc
#from clicknium.common.models.MouseLocation import MouseLocation
from SUBPROGRAMS.parameters import *
import pdfplumber


import time
import subprocess
import pyautogui
from datetime import datetime, timedelta
import re
import pygetwindow as gw
import os
import pyperclip
import yagmail
import yaml
import logging
import os
import sys
from datetime import datetime
from timeit import default_timer as timer
import json




'''
base_path = os.getcwd()
log_path = f'{base_path}/LOG'

image_path = f'{base_path}/IMAGES'
personalizar = f'{image_path}/personalizar.PNG'
valores_aba = f'{image_path}/valores_aba.PNG'
conta_caixa_botao = f'{image_path}/conta_caixa_botao.PNG'
descricao = f'{image_path}/descricao.PNG'
button_ok_3 = f'{image_path}/button_ok_3.PNG'
numero_documento = f'{image_path}/numero_documento.PNG'
barra_de_rolagem_direita = f'{image_path}/barra_rolagem_direita.PNG'
barra_de_rolagem_esquerda = f'{image_path}/barra_rolagem_esquerda.PNG'
imagem_natureza = f'{image_path}/imagem_natureza.PNG'

with open('config.yaml', 'r', encoding='utf=8') as params:
    config = yaml.safe_load(params)

print(config)

email_login = config['email']['login']
email_password = config['email']['password']
email_receivers = config['email']['receivers']
log_level = config['logLevel']

# log configs
today = datetime.today().strftime('%d-%m-%Y')
log_file_name = datetime.now().strftime('%H_%M_%S')

if  os.path.exists(f'{log_path}/{today}') == False:
    os.makedirs(f'{log_path}/{today}')

logging.basicConfig(level=log_level, datefmt='%d-%m-%Y %H:%M:%S',
                    format='%(asctime)s.%(msecs)03dZ;'
                               '%(module)s.%(funcName)s;'
                               '  %(message)s',
                    handlers=[
                        logging.FileHandler(os.path.join(f'{log_path}/{today}/{log_file_name}.log'), mode='w', encoding='utf-8', delay=False),
                        logging.StreamHandler(sys.stdout)
                    ])


# wait for image with pyautogui
def waitForImage(image, timeout, name):
    inicialTime = time.time()

    while True:
        try:
            # try to find image on screen
            location = pyautogui.locateOnScreen(image=image, confidence=0.9)

            if location is not None:
                # Image found
                logging.info(f'{name} image found.')
                return location

            
        except:
            # checking if timeout is over
            currentTime = time.time()
            if currentTime - inicialTime >= timeout:
                # timeout is over and the image was not found.
                logging.info(f'{name} image not found.')
                return None

            # waiting for searching again
            time.sleep(0.1)

def check_inv(vetor_inv, vetor_coligada):

    rows_int = 52 #################

    time.sleep(3)

    #SELECIONAR PRIMEIRA CÉLULA INDEPENDENTE DE ONDE ESTEJA
    cc.mouse.scroll(rows_int)
    referencia_inicial = cc.wait_appear(locator.rm.dataitem_x_row0, wait_timeout=360) 
    posicao = referencia_inicial.get_position()
    x, y = posicao.Right + 20 , (posicao.Bottom + posicao.Top) // 2
    print(x,y)
    pyautogui.moveTo(x, y)
    pyautogui.click()
    for esquerda in range (35):
        pyautogui.press('left')
    logging.debug('35 PASSOS PARA ESQUERDA')

    # PEGANDO DA PRIMEIRISSIMA CELULA
    for direita in range (19):
        pyautogui.press('right')
    logging.debug('19 PASSOS PARA DIREITA')

    time.sleep(2)
    for i in range(rows_int):

        logging.debug(f'BUSCA NA LINHA: {i}')


        #vetor_inv = informacoes_extraidas()
        padrao = re.compile(r'INV - (\d+) - (.+)$')

        logging.debug('ctrl c')
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.5)
        historico = pyperclip.paste()
        logging.debug(f'Historico completo: {historico}')

        localizar = padrao.search(historico)
        if localizar:
            inv_number = localizar.group(1)
            nome_coligada = localizar.group(2)
        else:
            inv_number = 'erro: INV NÃO ENCONTRADA'
            nome_coligada = 'erro: COLIGADA NÃO ENCONTRADA'

        #logging.debug(f'INV ENCONTRADA: {inv_number}')

        if inv_number in vetor_inv and vetor_coligada in nome_coligada:
            time.sleep(0.5)
            logging.debug(f'INV CORRETA ENCONTRADA: {inv_number}')
            logging.debug(f'COLIGADA ENCONTRADA: {nome_coligada}')
            break
            
        else:
            logging.debug(f'INV E COLIGADA INCORRETAS ENCONTRADAS: {inv_number} - {nome_coligada}')
            logging.debug(f'TESTE: {rows_int -1}')
            if i == rows_int - 1:
                erro = f'ULTIMA INV E COLIGADA CHECADAS E NÃO ENCONTRADAS- {inv_number}'
                logging.debug(f'{erro}')
                return erro
            time.sleep(0.5)
            pyautogui.press('down')
            logging.debug('APERTA PARA BAIXO PARA VERIFICAR O PRÓXIMO REGISTRO')
    

#check_inv('02023618', 'TSCHUDI SHIP MANAGEMENT AS')'''
            
def extrair_pdf():

    arquivo_extrato = r"\\Riofs\dados\Documentos Financeiro\Extratos bancarios\Ebury\2024\01 - Janeiro\Ebury - Completo.pdf"

    padroes_encontrados = []
    erro_invoice = []

    with pdfplumber.open(arquivo_extrato) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                # Procurar por ocorrências de "Pagamento PI" no texto da página
                ocorrencias = [match.start() for match in re.finditer(r'Pagamento PI', texto)]
                # Adicionar o número da página e a posição de cada ocorrência à lista de padrões encontrados
                for ocorrencia in ocorrencias:
                    data_match = re.search(r'(\d{2}-\d{2}-\d{4})', texto[ocorrencia:])
                    data = data_match.group(1) if data_match else None
                    
                    numero_pi_match = re.search(r'PI(\d+)', texto[ocorrencia:])
                    numero_pi = numero_pi_match.group(1) if numero_pi_match else None
                    
                    nome_empresa_match = re.search(r'Pagamento PI\d+ (.*?)\n', texto[ocorrencia:])
                    nome_empresa = nome_empresa_match.group(1) if nome_empresa_match else None
                    #nome_empresa = None
                    
                    #invoice_match = re.search(r'PI\d+\s+([\d./-]+)\b', texto[ocorrencia:])
                    invoice_match = re.search(r'PI\d+\s+(\d+)\b', texto[ocorrencia:])
                    invoice = invoice_match.group(1) if invoice_match else None
                    
                    moeda_match = re.search(r'\b[A-Z]{3}\b', texto[ocorrencia:])
                    moeda = moeda_match.group(0) if moeda_match else None
                    
                    '''montante_match = re.search(r'\b-?\d+\.\d{2}\b', texto[ocorrencia:])
                    montante = montante_match.group(0) if montante_match else None'''
                    
                    valores = re.search(r'[A-Z]{3}\b (-?\d+\.\d{2})\b\s+[A-Z]{3}\b (-?\d+\.\d{2})\b', texto[ocorrencia:])
                    #saldo = saldo_match.group(1) if saldo_match else None
                    montante = valores.group(1) if valores else None
                    saldo = valores.group(2) if valores else None

                    #padroes_encontrados.append((data, numero_pi, nome_empresa, invoice, moeda, montante, saldo))

                    # VALOR ESTRANGEIRO: MONTANTE, VALOR REAL: SALDO
                    padrao = {
                            "pi": {numero_pi},
                            "inv": {invoice},
                            "moeda": {moeda},
                            "montante": {montante},
                            "saldo": {saldo},
                            "nome": {nome_empresa},
                            "cod": None
                        }
                    
                    if invoice != None:
                        padroes_encontrados.append(padrao)

                    else:
                        erro_invoice.append(padrao)



    return padroes_encontrados, erro_invoice

#padroes_encontrados, erro_invoice = extrair_pdf()

'''for item in padroes_encontrados:
    print("certo: ",item)


for item in erro_invoice:
    print("errado: ",item)'''

#print(padroes_encontrados)

'''novo_filtro = waitForImage(image = novo_filtro_img, timeout=30, name = 'NOVO FILTRO')
if novo_filtro is not None:
    pyautogui.click(novo_filtro)
    logging.info('CLICK - BOTÃO DE NOVA SELECAO DE FILTRO')'''


'''def excelErro(erro_invoice):

    logging.info('Criando EXCEL ERRO')

    #workbook = openpyxl.Workbook()

    # Acesse a planilha ativa (por padrão, há uma planilha criada)
    sheet = workbook.active

    # Escreva dados nas células
    sheet['A1'] = 'DATA'
    sheet['B1'] = 'PI'
    sheet['C1'] = 'INV'
    sheet['D1'] = 'MOEDA'
    sheet['E1'] = 'VALOR'
    sheet['F1'] = 'NOME'
    sheet['G1'] = 'COD'
    sheet['H1'] = 'ERRO'
    sheet['I1'] = 'STATUS'

    logging.info('EXCEL ERRO iniciado')

    for i, extrato in enumerate(erro_invoice):

        logging.info(f'PROCESSO {i+1} A SER TRATADO: {extrato}')
        
        data = extrato["data"]
        for _ in range(2):
            data = data.replace("-", "")
        pi = extrato["pi"]
        inv = extrato["inv"]
        moeda = extrato["moeda"]
        montante = extrato["montante"]
        valor = extrato["saldo"]
        nome = extrato["nome"]
        cod = extrato["cod"]
        erro = 'PROCESSOS EXTRAÍDOS COM ERROS'
        status = 'NÃO FINALIZADO'

        logging.info(f'Criando linha do EXCEL ERRO: {pi} - {data} - {inv} - {moeda} - {valor} - {nome} - {cod} - {erro} - {status}')
        sheet = workbook.active

        sheet[f'A{i+2}'] = data
        sheet[f'B{i+2}'] = pi
        sheet[f'C{i+2}'] = inv
        sheet[f'D{i+2}'] = moeda
        sheet[f'E{i+2}'] = valor
        sheet[f'F{i+2}'] = nome
        sheet[f'G{i+2}'] = cod
        sheet[f'H{i+2}'] = erro
        sheet[f'I{i+2}'] = status

    logging.info(f'Fechando Excel com Erro...')
    workbook.save(f'{writer_erro}')  
'''

#lista1 = json.loads(informacoes_extraidas1())
'''
with open('teste.json', 'r') as arquivo1:
    # Carrega o conteúdo do arquivo JSON em um dicionário Python
    conteudo = json.load(arquivo1)
    print(conteudo)

lista = [
    {
        "pi": "",
        "inv": "000747431",
        "moeda": "NOK",
        "valor": "1000",
        "nome": "",
        "cod": ""
    },
    {
        "pi": "",
        "inv": "01401514",
        "moeda": "GBP",
        "valor": "2000",
        "nome": "",
        "cod": ""
    },
    {
        "data": "31-01-2024",
        "pi": "3836881",
        "inv": "31122023",
        "moeda": "EUR",
        "montante": "-440140.00",
        "saldo": "29580.64",
        "nome": "Tschudi Ship Management AS",
        "cod": None
    }
]

for i in lista:
    print(i)
    print(lista)
    if i in conteudo:
        print('sim')

    else:
        print('nao')

        conteudo.append(i)

        with open('teste.json', 'w') as arquivo:
            json.dump(conteudo, arquivo, indent=4)'''
    

import os
import glob



'''def excelProcessoCompleto(lista_processos):

    logging.info('Adicionando aba de PROCESSOS EXTRAÍDOS ao EXCEL existente')
    workbook = openpyxl.Workbook()

    # Criar uma nova aba no workbook existente
    sheet = workbook.create_sheet(title="PROCESSOS EXTRAÍDOS EXTRAÍDOS")

    # Escreva dados nas células
    sheet['A1'] = 'DATA'
    sheet['B1'] = 'PI'
    sheet['C1'] = 'INV'
    sheet['D1'] = 'MOEDA'
    sheet['E1'] = 'VALOR'
    sheet['F1'] = 'NOME'
    sheet['G1'] = 'COD'
    sheet['H1'] = 'ERRO'
    sheet['I1'] = 'STATUS'

    logging.info('Aba de ERRO adicionada ao EXCEL existente')

    for i, extrato in enumerate(lista_processos):
        logging.info(f'PROCESSO {i+1} A SER TRATADO: {extrato}')
        
        data = extrato["data"].replace("-", "")
        pi = extrato["pi"]
        inv = extrato["inv"]
        moeda = extrato["moeda"]
        valor = extrato["saldo"]
        nome = extrato["nome"]
        cod = extrato["cod"]
        erro = 'PROCESSOS EXTRAÍDOS COM ERROS'
        status = 'NÃO FINALIZADO'

        # Escrever dados na nova aba
        sheet[f'A{i+2}'] = data
        sheet[f'B{i+2}'] = pi
        sheet[f'C{i+2}'] = inv
        sheet[f'D{i+2}'] = moeda
        sheet[f'E{i+2}'] = valor
        sheet[f'F{i+2}'] = nome
        sheet[f'G{i+2}'] = cod
        sheet[f'H{i+2}'] = erro
        sheet[f'I{i+2}'] = status

    #logging.info('Fechando Excel com Erro...')
    workbook.save(f'{writer_erro}')
        

# Defina o caminho da pasta que deseja buscar os arquivos PDF
pasta = r'\\Riofs\dados\Documentos Financeiro\Extratos bancarios\Ebury\2024\01 - Janeiro'
#pasta = 'Y:\01 - Janeiro'

# Use a função glob para buscar todos os arquivos PDF na pasta especificada
arquivos_pdf = glob.glob(os.path.join(pasta, '*.pdf'))
print(arquivos_pdf)

# Exiba a lista de arquivos PDF encontrados
for arquivo in arquivos_pdf:

    excelProcessoCompleto(arquivo)
    print(arquivo)'''



'''def process():
    global arquivo

    locais = [r"\\Riofs\dados\Documentos Financeiro\Extratos bancarios\Ebury\2024\01 - Janeiro"]

    for pas, local in enumerate(locais):

        inicial()


        arquivos_pdf = glob.glob(os.path.join(local, '*.pdf'))
        logging.info(f'SERÃO TRATADOS {len(arquivos_pdf)} ARQUIVOS.')


        for arq, arquivo in enumerate(arquivos_pdf):
            
            logging.info(f'ARQUIVO A SER TRATADO: {arquivo}')

                
            #erro_invoice = ''
            #extrato_list_nao_tratado = informacoes_extraidas()
            #extrato_list = json.loads(extrato_list_nao_tratado)
            
            
            extrato_list_nao_tratado, erro_invoice = informacoes_extraidas(arquivo)

            logging.info(f'PROCESSOS EXTRAIDOS COM ERRO: {len(erro_invoice)}')
            #logging.info(f'PROCESSOS COM ERRO: {erro_invoice}')
            excelErro(erro_invoice)

            logging.info(f'PROCESSOS EXTRAIDOS CORRETAMENTE: {len(extrato_list_nao_tratado)}')
            ##logging.info(f'PROCESSOS A SEREM TRATADOS: {extrato_list_nao_tratado}')

            #logging.info(f'PROCESSOS EXTRAIDOS CORRETAMENTE: {len(extrato_list)}')
            #logging.info(f'PROCESSOS A SEREM TRATADOS: {extrato_list}')

            for i, extrato in enumerate(extrato_list_nao_tratado):
            #for i, extrato in enumerate(extrato_list):

                if i != 0:
                    pass
                    #logging.info(f'PASSANDO PARA PROXIMO PROCESSO')


                #logging.info(f'PROCESSO {i+1} A SER TRATADO: {extrato}')

                

                data = extrato["data"]
                for _ in range(2):
                    data = data.replace("-", "")

                pi = extrato["pi"]

                inv_extraida = extrato["inv"]
                #logging.info(f'inv original {inv_extraida}. tamanho {len(inv_extraida)}')
                inv = f"{inv_extraida}01"
                while len(inv) < 11:
                    inv = f"0{inv}"
                #logging.info(f'inv TRATADA {inv}. tamanho {len(inv)}')

                moeda = extrato["moeda"]

                valor = extrato["montante"]
                valor = valor.replace("-", "")
                valor = valor.replace(".", ",")

                saldo = extrato["saldo"]
                nome = extrato["nome"]
                cod = extrato["cod"]
                
                erro = f'{arq}{pas}'
                status = f'{arq}{pas}'

                createRowExcel(i, data, pi, inv_extraida, moeda, valor, nome, cod, erro, status) 

        encerra_prog()

'''
#process()
        

data_atual = datetime.now()
mes_atual = data_atual.month
logging.info(type(mes_atual))

check_mes = verificar_arquivo_vazio()
if check_mes != True:
    mes_atual = ler_arquivo()
    logging.info(type(mes_atual))


#pasta_local = f'{pasta_original}\\{config['mes'][mes_atual]}'
#logging.info(f'Pasta local: {pasta_local}')