from SUBPROGRAMS.parameters import *

from clicknium import locator, ui
from clicknium import clicknium as cc

def bot_text_art():
    logging.info("    ____   ")
    logging.info("   [____]    ")
    logging.info(" |=]()()[=|  ")
    logging.info("   _\==/__    _____    ____     _        _____    ___   _  __")
    logging.info(" |__|   |_|  |  _  \  / ___|   | |      |_   _| / ___| | |/ /")
    logging.info(" |_|_/\_|_|  | |  | | | |      | |        | |   | |    | ' / ")
    logging.info(" | | __ | |  | |  | | | |      | |        | |   | |    |  <  ")
    logging.info(" |_|[  ]|_|  | |__| | | |____  | |____   _| |_  | |__  | . \ ")
    logging.info(" \_|_||_|_/  |_____/   \_____| |______| |_____| \ ___| |_|\_\\")
    logging.info("   |_||_|                                                     ")
    logging.info("   | ||_|_   ")
    logging.info(" |___||___|  ")	
    logging.info("             ")


def show_exception_and_exit(exc_type, exc_value, tb):

    logging.error(exc_value, exc_info=(exc_type, exc_value, tb))

    result = datetime.now().strftime('%Y_%m_%d %H_%M')
    erro_print = f'ERRO {result}'

    screenshot = pyautogui.screenshot()
    screenshot.save(f"{print_path}\\{erro_print}.png")
    logging.info(f'PRINT SALVO')
    sleep(1)

    encerra_prog()

    with open(f'{email_path}/emailError.html', 'r', encoding='utf-8') as fileError:
        templateError = fileError.read()
        htmlError = templateError.format(mensagem=exc_value)
        
    yag.send(to=email_receivers.split(','), subject=f"TESTES - ERRO EXCEÇÃO - BOT RETORNO PAGAMENTO EM MOEDA ESTRANGEIRA - BV",
    contents=htmlError,
    attachments=[f'{log_path}/{today}/{log_file_name}.log', f'{print_path}/{erro_print}.png', writer])

    sys.exit(-1)


def open_RM():

    subprocess.Popen(localRM)
    logging.info('OPEN PROGRAM RM')
    

def loggin_RM():

    cc.wait_appear(locator.rm.pane_picturebox1, wait_timeout=1200)

    '''if 'capslock' in pyautogui.KEYBOARD_KEYS:
        pyautogui.press('capslock')  # Pressiona Caps Lock para alternar seu estado
        logging.info('capslock CLICADO')
        sleep(0.1) '''

    user_rm = cc.wait_appear(locator.rm.user_field, wait_timeout=15)
    user_rm.double_click()
    pyautogui.write(f'{user}')
    logging.info('ENTER USER_NAME')

    pass_rm = cc.wait_appear(locator.rm.pass_field, wait_timeout=15)
    pass_rm.click()
    pyautogui.typewrite(f'{password}')
    logging.info('ENTER PASSWORD')

    click_button = cc.wait_appear(locator.rm.button_entrar, wait_timeout=15)
    click_button.click()
    logging.info('ENTER RM')

    licencas = waitForImage(image = licencas_excedidas, timeout=40, name = 'LICENCAS EXCEDIDAS')
    if licencas is not None:
        logging.info('LICENCAS EXCEDIDAS')
        sleep(1)
        pyautogui.hotkey(f'esc')

        licencas2 = waitForImage(image = licencas_excedidas2, timeout=20, name = 'LICENCAS EXCEDIDAS 2')
        if licencas2 is not None:
            logging.info('LICENCAS EXCEDIDAS erro 2')
            sleep(1)
            pyautogui.hotkey(f'esc')

    



def enter_coligada():

    click_context = cc.wait_appear(locator.rm.click_contexto, wait_timeout=300)
    click_context.click()
    logging.info('CLICK CONTEXTO')
    click_coligada = cc.wait_appear(locator.rm.menuitem_alterar_contexto_do_módulo, wait_timeout=15)
    click_coligada.click()
    logging.info('CLICK COLIGADA 2')
    click_colig_cod = cc.wait_appear(locator.rm.edit_coligada_code, wait_timeout=15)
    click_colig_cod.click()
    click_colig_cod.set_text('2')
    logging.info('INSERE O COLIGADA CODE = 2')
    pyautogui.press('tab')
    logging.info('APERTA TAB PARA PREENCHER O CAMPO COM A STRING: OCEANICA ENGENHARIA E CONSULTORIA S.A.')
    click_avancar_final = cc.wait_appear(locator.rm.button_bavancar1, wait_timeout=15)
    click_avancar_final.click()
    logging.info('CLICK AVANÇAR NOVAMENTE')
    click_concluir = cc.wait_appear(locator.rm.button_concluir, wait_timeout=15)
    click_concluir.click()
    sleep(2)
    logging.info('CLICK CONCLUIR')
    try:
        click_concluir = cc.wait_disappear(locator.rm.button_concluir, wait_timeout=120)
        logging.info('POP UP COLIGADA CLICADO PARA DESAPARECER')
    except:
        logging.info('POP UP COLIGADA NÃO EXIBIDO')
        pass



def enter_contas_a_pagar():

    linha_totvs = cc.wait_appear(locator.rm.linha_rm, wait_timeout=120)
    linha_totvs.click()
    logging.info('CLICK - TOTVS - LINHA RM')
    sleep(2)

    back_office = cc.wait_appear(locator.rm.back_office, wait_timeout=30)
    back_office.click()
    logging.info('CLICK - BACK OFFICE')
    sleep(2)

    gestao_financeira = cc.wait_appear(locator.rm.gestao_financeira, wait_timeout=30)
    gestao_financeira.click()
    logging.info('CLICK - GESTAO FINANCEIRA')
    sleep(2)

    contas_pagar = cc.wait_appear(locator.rm.contas_pagar, wait_timeout=60)
    contas_pagar.click()
    logging.info('CLICK - CONTAS A PAGAR')

    '''select_lancamentos = waitForImage(image = lancamentos, timeout=60, name = 'LANÇAMENTOS')
    if select_lancamentos is not None:
        #sleep(2)
        pyautogui.click(select_lancamentos)
        logging.info('CLICA EM LANÇAMENTOS')
        sleep(2)'''

    cont = 0
    while cont <= 50:

        select_lancamentos = waitForImage(image = lancamentos, timeout=10, name = 'LANÇAMENTOS')
        if select_lancamentos is not None:
            pyautogui.click(select_lancamentos)
            logging.info('CLICA EM LANÇAMENTOS')
            sleep(5)

        licencas = waitForImage(image = licencas_excedidas, timeout=20, name = 'LICENCAS EXCEDIDAS')
        if licencas is not None:
            logging.info('LICENCAS EXCEDIDAS')
            logging.info(f'TENTANDO NOVAMENTE - {cont+1}')
            sleep(2)
            pyautogui.hotkey(f'esc')
            if cont == 50:
                texto_msg = 'LIMITE DE TENTATIVAS PARA ADQUIRIR UMA LICENÇA ALCANÇADO'
                logging.info(texto_msg)
                encerra_prog()
                emailErro(texto_msg)
            cont += 1
            sleep(20)
        
        else:
            break


def obter_intervalo_data():
    hoje = datetime.now()
    dia_semana_hoje = hoje.weekday()  # Retorna um número entre 0 (segunda) e 6 (domingo)

    if dia_semana_hoje in [2, 3, 4, 5, 6]:  # Quarta, quinta ou sexta
        data_inicio = hoje - timedelta(days=dia_semana_hoje - 2)  # Início da quarta-feira atual
        data_fim = data_inicio + timedelta(days=6)  # Fim da terça-feira da próxima semana
    elif dia_semana_hoje in [0, 1]:  # Segunda ou terça
        data_inicio = hoje - timedelta(days=dia_semana_hoje + 5)  # Início da quarta-feira passada
        data_fim = data_inicio + timedelta(days=6)  # Fim da terça-feira desta semana
    else:
        # Lidar com outros dias da semana conforme necessário
        data_inicio = data_fim = None

    return data_inicio, data_fim

def filter_lancamento(invoice, i, cont, flag):

    if i == 0 and cont == 0 or flag:
        logging.info('filter_lancamento : 1')

        '''select_filter = cc.wait_appear(locator.rm.select_filter, wait_timeout=60)
        sleep(2)
        select_filter.click()
        #iltro = 'histórico'
        #filtro = 'Valor Original'
        filtro = 'Número do documento'
        sleep(4)

        select_filter.set_text(f'{filtro}')
        logging.info(f'ESCREVE {filtro} NO FILTRO')
        sleep(1)'''
                
        select_filter = cc.wait_appear(locator.rm.select_filter, wait_timeout=60)
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
        logging.info('CLICK - BOTÃO DE SELECAO DE FILTRO')

    else:
        logging.info('filter_lancamento : 2')

        cc.wait_appear(locator.rm.checkbox_lançamentos_a_receber, wait_timeout=360)

        sleep(1)

        novo_filtro = waitForImage(image = novo_filtro_img, timeout=30, name = 'NOVO FILTRO')
        if novo_filtro is not None:
            pyautogui.click(novo_filtro)
            logging.info('CLICK - BOTÃO DE NOVA SELECAO DE FILTRO')

        '''execute = cc.wait_appear(locator.rm.splitbutton_filtro_histórico, wait_timeout=15)
        execute.click()
        logging.info('CLICK - BOTÃO DE NOVA SELECAO DE FILTRO')'''


    enter_inv = cc.wait_appear(locator.rm.edit, wait_timeout=15)
    enter_inv.double_click()
    sleep(1)

    #pyautogui.typewrite(f'INV')
    pyautogui.typewrite(invoice)
    logging.info(f'VALOR DE {invoice} FILTRO')


    ok_filter = cc.wait_appear(locator.rm.button_ok, wait_timeout=15)
    ok_filter.click()
    logging.info('CLICK - BOTÃO DE EXECUTAR O FILTRO')


def filter_clean():
    for i in range(3):
        clear_filter_x = waitForImage(image = filter_clear, timeout=10, name = 'FILTER CLEAN')
        if clear_filter_x is not None:
            pyautogui.click(clear_filter_x)
            logging.info(f'CLICA NO BOTÃO X DO LIMPA FILTRO - TENTATIVA {i+1}')   
            sleep(2)
        else:
            logging.info(f'BOTÃO X DO LIMPA FILTRO NÃO ENCONTRADO - TENTATIVA {i+1}')   


def select_tipo_documento():

    tipo_doc = cc.wait_appear(locator.rm.header_tipo_de_documento, wait_timeout=600)
    colunm_name = tipo_doc.get_text()
    logging.info(f'NOME DA COLUNA: {colunm_name}')

    tipo_doc_hover = cc.wait_appear(locator.rm.header_tipo_de_documento, wait_timeout=15)
    tipo_doc_hover.hover(4)
    logging.info(f'HOVER NOME DA COLUNA: {colunm_name}')

    posição = tipo_doc_hover.get_position()
    x, y = posição.Right - 5 , posição.Top + 5
    pyautogui.moveTo(x,y)
    pyautogui.click()
    sleep(1)

    pyautogui.press('down')

    pyautogui.typewrite('(Personalizar)')
    pyautogui.press('enter')
    logging.info('CLICA NO FILTRO PERSONALIZAR')

    into_filter_personalizar = cc.wait_appear(locator.rm.listitem_filter_type, wait_timeout=15)
    into_filter_personalizar.click()
    logging.info('CLICA PARA SELECIONAR OPÇÕES DO FILTRO')

    sleep(2)
    pyautogui.write('Igual a')
    sleep(2)
    pyautogui.hotkey('tab')
    sleep(2)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.write(F'{inv_palavra}')

    button_not_inv = cc.wait_appear(locator.rm.button_confirmar_not_inv, wait_timeout=15)
    button_not_inv.click()
    logging.info('CLICK BUTTON CONFIRMAR FILTRO')


def desmark_filter():

    check_lancamentos_receber = cc.wait_appear(locator.rm.checkbox_lançamentos_a_receber, wait_timeout=360)
    check_lancamentos_receber.click()
    logging.info('CLICK CHECKBOX LANÇAMENTOS A RECEBER')

    check_lancamentos_baixados = cc.wait_appear(locator.rm.checkbox_lançamentos_baixados, wait_timeout=15)
    check_lancamentos_baixados.click()
    logging.info('CLICK CHECKBOX LANÇAMENTOS BAIXADOS')

    check_lancamentos_cancelados = cc.wait_appear(locator.rm.checkbox_lançamentos_cancelados, wait_timeout=15)
    check_lancamentos_cancelados.click()
    logging.info('CLICK CHECKBOX LANÇAMENTOS CANCELADOS')

    check_lancamentos_faturados = cc.wait_appear(locator.rm.checkbox_lançamentos_faturados, wait_timeout=15)
    check_lancamentos_faturados.click()
    logging.info('CLICK CHECKBOX LANÇAMENTOS FATURADOS')

    filter_refresh = cc.wait_appear(locator.rm.button_atualiza_filtro, wait_timeout=15)
    filter_refresh.click()
    logging.info('ATUALIZA O FILTRO')


def check_table_record():
    global rows_int
    logging.info('VERIFICANDO SE A TABELA ESTÁ VAZIA')

    cc.wait_appear(locator.rm.checkbox_lançamentos_a_receber, wait_timeout=360)

    logging.info('VERIFICAdo')

    #sleep(10)  AGUARDAR A CONFIRMAÇÃO DE UMA LINHA APARECER
    cc.wait_appear(locator.rm.dataitem_x_row0, wait_timeout=5)

    pyautogui.hotkey('ctrl', 'g')
    logging.info('CTRL+G')   

    numero_registros = cc.wait_appear(locator.rm.text_table_row, wait_timeout=15)
    txt_table = numero_registros.get_text()
    logging.info(txt_table)
    sleep(2)
    number_rows = re.findall(r'\b(\d+)\b', txt_table)

    if txt_table == 'Para este processo é necessário que existam registros na visão.':
        
        sleep(1)
        logging.info('SEM REGISTROS ENCONTRADOS')   
        pyautogui.press('enter')
        return True
    
    else:

        rows_int = int(number_rows[0])
        logging.info(f'NUMERO DE REGISTROS ENCONTRADOS: {rows_int}')
        #sleep(5)
        click_nao = cc.wait_appear(locator.rm.button_não, wait_timeout=15)
        click_nao.set_focus()
        sleep(1)
        pyautogui.hotkey('alt', 'n')
        logging.info('CLICK NÃO')
        return False 
    
        windows = gw.getWindowsWithTitle(title='TOTVS Linha RM - Construção e Projetos  Alias: CorporeRM | 1-OCEANICA ENGENHARIA E CONSULTORIA S.A.')
        print(windows)
        if windows:
            window=windows[0]
            window.activate()
            window.maximize()


def encerra_prog():
        
        logging.info(f'Fechando Excel...')
        #workbook.save(f'{writer}')  
        workbook.save(f'{writer}')  
        
        pidKillFinish()

        logging.info('RM ENCERRADO')

'''def check_inv():

    sleep(5)
    select_first_row = cc.wait_appear(locator.rm.dataitem_ref_lançamento_row0, wait_timeout=120)
    select_first_row.click()
    logging.info('SELECIONA A PRIMEIRA LINHA REF LANÇAMENTO')

    for i in range (16):
        pyautogui.press('right')
    logging.info('16 PASSOS PARA DIREITA')

    for i in range(rows_int):

        vetor_inv = []

        sleep(2)
        logging.info('ctrl c')
        pyautogui.hotkey('ctrl', 'c')
        sleep(2)
        inv_number = pyperclip.paste()
        logging.info(f'INV: {inv_number}')

        if inv_number in vetor_inv:
            sleep(2)
            pyautogui.hotkey('ctrl', 'enter')

            select_valores = waitForImage(image = valor, timeout=30, name = 'INTEGRAÇÃO BANCÁRIA')
            if select_valores is not None:
                sleep(2)
                pyautogui.click(select_valores)
                logging.info('CLICK EM VALOR')

            for i in range (4):
                pyautogui.press('tab')
            logging.info('4 TAB')

            pyautogui.press('enter')
            sleep(3)
            pyautogui.write(vetor_inv.moeda)
            sleep(2)
            pyautogui.press('enter')
            sleep(2)
            pyautogui.press('enter')

            btt_ok = waitForImage(image = button_ok, timeout=10, name = 'PRESS BUTTON OK')
            if btt_ok is not None:
                sleep(2)
                pyautogui.click(btt_ok)
                logging.info('APERTA OK')
                sleep(2)

            baixa_inv()
            
        else:
            pyautogui.press('down')
            logging.info('APERTA PARA BAIXO PARA VERIFICAR O PRÓXIMO REGISTRO')
            sleep(2)'''


def check_inv(vetor_inv, vetor_coligada, valor_extrato):

    #rows_int = 52 #################

    cc.wait_appear(locator.rm.checkbox_lançamentos_a_receber, wait_timeout=360)

    largura_tela, altura_tela = pyautogui.size()

    # Calcula as coordenadas do meio da tela
    meio_x = largura_tela // 2
    meio_y = altura_tela // 2

    pyautogui.moveTo(meio_x, meio_y)

    sleep(1)

    #SELECIONAR PRIMEIRA CÉLULA INDEPENDENTE DE ONDE ESTEJA
    cc.mouse.scroll(rows_int)
    logging.info(f'MOVIDO PRA CIMA {rows_int} VEZES')

    referencia_inicial = cc.wait_appear(locator.rm.dataitem_x_row0, wait_timeout=360) 
    posicao = referencia_inicial.get_position()
    x, y = posicao.Right + 20 , (posicao.Bottom + posicao.Top) // 2
    pyautogui.moveTo(x, y)
    pyautogui.click()
    logging.info('PRIMEIRA CELULA CLICADA')

    for esquerda in range (35):
        pyautogui.press('left')
    logging.info('35 PASSOS PARA ESQUERDA')

    # PEGANDO DA PRIMEIRISSIMA CELULA
    for direita in range (18):
        pyautogui.press('right')
    logging.info('18 PASSOS PARA DIREITA')

    sleep(2)

    for i in range(rows_int):

        logging.info(f'BUSCA NA LINHA: {i+1}')


        #vetor_inv = informacoes_extraidas()
        padrao = re.compile(r'INV - (\d+) - (.+)$')

        logging.info('ctrl c')
        pyautogui.hotkey('ctrl', 'c')
        sleep(0.5)
        historico = pyperclip.paste()
        logging.info(f'Buscando: {vetor_inv} - {vetor_coligada}')
        logging.info(f'Historico completo encontrado: {historico}')

        localizar = padrao.search(historico)
        if localizar:
            inv_number = localizar.group(1)
            nome_coligada = localizar.group(2)
        else:
            inv_number = 'erro: INV NÃO ENCONTRADA'
            nome_coligada = 'erro: COLIGADA NÃO ENCONTRADA'

        #logging.info(f'INV ENCONTRADA: {inv_number}')

        inv_number1 = inv_number.lower()
        vetor_inv1 = vetor_inv.lower()
        vetor_coligada1 = vetor_coligada.lower()
        nome_coligada1 = nome_coligada.lower()

        vetor_coligada2 = vetor_coligada1.replace('.','')
        vetor_coligada21 = vetor_coligada2.replace(' - ',' ')
        vetor_coligada22 = vetor_coligada21.replace(' e ',' & ')
        vetor_coligada23 = vetor_coligada22.replace(' and ',' & ')
        vetor_coligada24 = vetor_coligada23.replace(' inc','')
        vetor_coligada25 = vetor_coligada24.replace(',','')
        if vetor_coligada25.endswith('s'):
            vetor_coligada26 = vetor_coligada25[:-1]
        else:
            vetor_coligada26 = vetor_coligada25

        vetor_coligada3 = vetor_coligada1.replace(' - ',' ')
        vetor_coligada31 = vetor_coligada3.replace('.','')
        vetor_coligada32 = vetor_coligada31.replace(' e ',' & ')
        vetor_coligada33 = vetor_coligada32.replace(' and ',' & ')
        vetor_coligada34 = vetor_coligada33.replace(' inc','')
        vetor_coligada35 = vetor_coligada34.replace(',','')
        if vetor_coligada25.endswith('s'):
            vetor_coligada36 = vetor_coligada35[:-1]
        else:
            vetor_coligada36 = vetor_coligada35

        vetor_coligada4 = vetor_coligada1.replace(' e ',' & ')
        vetor_coligada41 = vetor_coligada4.replace('.','')
        vetor_coligada42 = vetor_coligada41.replace(' - ',' ')

        vetor_coligada5 = vetor_coligada1.replace(' and ',' & ')
        vetor_coligada51 = vetor_coligada5.replace('.','')
        vetor_coligada52 = vetor_coligada51.replace(' - ',' ')

        if inv_number1 in vetor_inv1 and vetor_coligada1 in nome_coligada1 or \
            \
            inv_number1 in vetor_inv1 and vetor_coligada2 in nome_coligada1 or \
            inv_number1 in vetor_inv1 and vetor_coligada21 in nome_coligada1 or \
            inv_number1 in vetor_inv1 and vetor_coligada22 in nome_coligada1 or \
            inv_number1 in vetor_inv1 and vetor_coligada23 in nome_coligada1 or \
            inv_number1 in vetor_inv1 and vetor_coligada24 in nome_coligada1 or \
            inv_number1 in vetor_inv1 and vetor_coligada25 in nome_coligada1 or \
            inv_number1 in vetor_inv1 and vetor_coligada26 in nome_coligada1 or \
            \
            inv_number1 in vetor_inv1 and vetor_coligada3 in nome_coligada1 or \
            inv_number1 in vetor_inv1 and vetor_coligada31 in nome_coligada1 or \
            inv_number1 in vetor_inv1 and vetor_coligada32 in nome_coligada1 or \
            inv_number1 in vetor_inv1 and vetor_coligada33 in nome_coligada1 or \
            inv_number1 in vetor_inv1 and vetor_coligada34 in nome_coligada1 or \
            inv_number1 in vetor_inv1 and vetor_coligada36 in nome_coligada1 or \
            \
            inv_number1 in vetor_inv1 and vetor_coligada4 in nome_coligada1 or \
            inv_number1 in vetor_inv1 and vetor_coligada41 in nome_coligada1 or \
            inv_number1 in vetor_inv1 and vetor_coligada42 in nome_coligada1 or \
            \
            inv_number1 in vetor_inv1 and vetor_coligada5 in nome_coligada1 or \
            inv_number1 in vetor_inv1 and vetor_coligada51 in nome_coligada1 or \
            inv_number1 in vetor_inv1 and vetor_coligada52 in nome_coligada1:


            sleep(0.5)
            logging.info(f'INV E COLIGADA CORRETA ENCONTRADA: {inv_number} - {nome_coligada}')

            for esquerda in range (4):
                pyautogui.press('left')
            logging.info('4 PASSOS PARA ESQUERDA')

            pyautogui.hotkey('ctrl', 'c')
            logging.info('ctrl c')
            sleep(0.5)

            valor_processo = pyperclip.paste()
            for _ in range(3):
                valor_processo = valor_processo.replace(".","")
            logging.info(f'VALOR PROCESSO: {valor_processo}')
            logging.info(f'VALOR PROCESSO2: {type(valor_processo)}')

            if valor_processo == valor_extrato:
                logging.info('VALOR CORRETO ENCONTRADO.')
                break
            
            
        if i+1 == rows_int:
            logging.info(f'ULTIMA INV, COLIGADA OU VALOR INCORRETAS ENCONTRADAS: {inv_number} - {nome_coligada}')
            
            erro = f'Processo não encontrado'
            logging.info(f'{erro}')
            return erro
        
        logging.info(f'INV, COLIGADA OU VALOR INCORRETAS ENCONTRADAS: {inv_number} - {nome_coligada}')
        sleep(0.3)

        pyautogui.press('down')
        logging.info('APERTA PARA BAIXO PARA VERIFICAR O PRÓXIMO REGISTRO')



def define_moeda(moeda):

    sleep(2)
    pyautogui.hotkey('ctrl','enter')
    logging.info('ctrl + enter')
    
    select_valores = waitForImage(image = valor, timeout=30, name = 'INTEGRAÇÃO BANCÁRIA')
    if select_valores is not None:
        sleep(2)
        pyautogui.click(select_valores)
        logging.info('CLICK EM VALOR')

    for i in range (4):
        pyautogui.press('tab')
    logging.info('4 TAB')

    sleep(30)

    pyautogui.press('enter')
    sleep(3)
    pyautogui.write(moeda)
    logging.info(f'MOEDA PREENCHIDA COM {moeda}')

    sleep(2)
    pyautogui.press('enter')
    sleep(2)
    pyautogui.press('enter')


    btt_ok = waitForImage(image = button_ok, timeout=10, name = 'PRESS BUTTON OK')
    if btt_ok is not None:
        sleep(2)
        pyautogui.click(btt_ok)
        logging.info('APERTA OK')


def informacoes_extraidas(arquivo_extrato):
    
    todos_processos = []
    padroes_encontrados = []
    erro_invoice = []

    with pdfplumber.open(arquivo_extrato) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                logging.info(f'texto: {texto}')
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
                    
                    #moeda_match = re.search(r'\b[A-Z]{3}\b', texto[ocorrencia:])
                    moeda_match = re.search(r'([A-Z]{3}) \d+.\d{2}', texto[ocorrencia:])
                    moeda = moeda_match.group(1) if moeda_match else None
                    
                    '''montante_match = re.search(r'\b-?\d+\.\d{2}\b', texto[ocorrencia:])
                    montante = montante_match.group(0) if montante_match else None'''
                    
                    valores = re.search(r'[A-Z]{3}\b (-?\d+\.\d{2})\b\s+[A-Z]{3}\b (-?\d+\.\d{2})\b', texto[ocorrencia:])
                    #saldo = saldo_match.group(1) if saldo_match else None
                    montante = valores.group(1) if valores else None
                    saldo = valores.group(2) if valores else None

                    #padroes_encontrados.append((data, numero_pi, nome_empresa, invoice, moeda, montante, saldo))

                    # VALOR ESTRANGEIRO: MONTANTE, VALOR REAL: SALDO
                    padrao = {
                            "data": data,
                            "pi": numero_pi,
                            "inv": invoice,
                            "moeda": moeda,
                            "montante": montante,
                            "saldo": saldo,
                            "nome": nome_empresa,
                            "cod": None
                        }
                    
                    todos_processos.append(padrao)

                    padrao_moeda = ['NOK','EUR','USD','GBP','BRL','SGD']
                    
                    if invoice != None and moeda in padrao_moeda:
                        padroes_encontrados.append(padrao)

                    else:
                        erro_invoice.append(padrao)

    excelProcessoCompleto(todos_processos, arquivo_extrato)



    return padroes_encontrados, erro_invoice
    
def informacoes_extraidas1():
    return"""
[
    {
        "data": "28-01-2024",
        "pi": "teste",
        "inv": "8002183",
        "moeda": "USD",
        "montante": "429.866,00",
        "saldo": "429.866,00",
        "nome": "KONGSBERG MARITIME AS",
        "cod": "teste2"
    },
    {
        "data": "28-01-2024",
        "pi": "teste",
        "inv": "2022685",
        "moeda": "EUR",
        "montante": "6.015,00",
        "saldo": "6.015,00",
        "nome": "TSCHUDI SHIP MANAGEMENT AS",
        "cod": "teste2"
    },
    {
        "data": "31-01-2024",
        "pi": "3836881",
        "inv": "31122023",
        "moeda": "EUR",
        "montante": "-440140.00",
        "saldo": "29580.64",
        "nome": "Tschudi Ship Management AS",
        "cod": null
    }
]
    """


"""
[
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
    """


def baixa_inv(valor_original, conta_certa, data_extrato):

    sleep(2)
    pyautogui.hotkey('ctrl', 'space')
    logging.info('CHECK LINHA')

    baixa_inv = cc.wait_appear(locator.rm.button_baixa, wait_timeout=60)
    baixa_inv.click()
    logging.info('SELECIONA BOTÃO BAIXA')

    try:
        button_ok = cc.wait_appear(locator.rm.button_btnok1, wait_timeout=30)
        button_ok.click()
        logging.info('CLICA OK')
    except:
        logging.info('POPUP BOTÃO OK NÃO APARECEU')
        pass

    avancar = cc.wait_appear(locator.rm.button_btnnext, wait_timeout=120)
    avancar.click()
    logging.info('CLICA AVANÇAR')

    sleep(5)

    simplificada = cc.wait_appear(locator.rm.radiobutton_simplificada, wait_timeout=120)
    simplificada.click()
    logging.info('SELECIONA SIMPLIFICADA')

    data = cc.wait_appear(locator.rm.edit_data, wait_timeout=60)
    date = data.get_text()
    logging.info(f'DATA AUTOMÁTICA: {date}')

    for _ in range (2):
        date = date.replace("/","")

    ## SE NECESSÁRIO DIFITAR DATA
    #data.set_text('20112023')
        
    if date != data_extrato:
        logging.info(f'DATAS DIFERENTES, FAZENDO ALTERAÇÃO')

        data.set_text(data_extrato)
        pyautogui.hotkey('tab')

    ###entrar if para compara com a data do documento inicial

    '''conta_caixa = cc.wait_appear(locator.rm.edit_lkpcontacaixacode, wait_timeout=60)
    conta = conta_caixa.get_text()'''

    
    erro1 = cc.wait_appear(locator.rm.titlebar_titlebar, wait_timeout=2)
    erro2 = cc.wait_appear(locator.rm.window_rm, wait_timeout=1)
    erro3 = cc.wait_appear(locator.rm.pane_rmspictureedit1, wait_timeout=1)

    if erro1 or erro2 or erro3:
        erro = f'ERRO ENVOLVENDO MOEDA OU DATA'
        logging.info(f'{erro}')
        pyautogui.hotkey('alt', 'f4')
        sleep(0.5)
        pyautogui.hotkey('alt', 'f4')
        return erro
    

    conta_caixa = cc.wait_appear(locator.rm.edit_lkpcontacaixalookupeditinneredit, wait_timeout=60)
    conta = conta_caixa.get_text()
    logging.info(f'CONTA CAIXA: {conta}')

    if conta_certa.lower() in conta.lower():
        logging.info('CONTA CAIXA PREENCHIDA CORRETAMENTE')
    else:
        logging.info('CONTA CAIXA NÃO PREENCHIDA CORRETAMENTE')
    
    
    
    avancar = cc.wait_appear(locator.rm.button_btnnext, wait_timeout=60)
    avancar.click()
    logging.info(f'AVANÇAR')

    valor = cc.wait_appear(locator.rm.edit_cvalordabaixadolancto, wait_timeout=60)
    valor_inv = valor.get_text()
    for _ in range(3):
        valor_inv = valor_inv.replace(".","")
    logging.info(f'VALOR INV: {valor_inv}')

    if valor_original != valor_inv:
        erro = f'VALOR DE DOCUMENTO DIFERENTE DO EXTRAIDO'
        logging.info(f'{erro}')
        for _ in range(1):
            sleep(2)
            pyautogui.hotkey('alt', 'f4')
            logging.info(f'Alt + f4 CLICADO')
        return erro

    avancar = cc.wait_appear(locator.rm.button_btnnext, wait_timeout=60)
    avancar.click()
    logging.info(f'AVANÇAR')

    avancar = cc.wait_appear(locator.rm.button_btnnext, wait_timeout=60)
    avancar.click()
    logging.info(f'AVANÇAR')

    exec = cc.wait_appear(locator.rm.button_btnnext, wait_timeout=60)
    exec.click()
    logging.info(f'EXECUTAR')

    processa_success = waitForImage(image = baixa_success, timeout=120, name = 'PRESS BUTTON OK')
    if processa_success is not None:
        sleep(2)
        logging.info('OK SUCESSO')
        sleep(2)
    else:
        erro = 'ERRO AO BAIXAR INV'
        logging.info(f'{erro}')
        pyautogui.hotkey('alt', 'f4')
        sleep(0.5)
        pyautogui.hotkey('alt', 'f4')
        return erro

    log = cc.wait_appear(locator.rm.edit_textboxlog3, wait_timeout=180)
    txt_log = log.get_text()
    logging.info(f'{txt_log}')

    close = cc.wait_appear(locator.rm.button_btncancel, wait_timeout=60)
    close.click()
    logging.info(f'FECHAR')



# wait for image with pyautogui
'''def waitForImage(image, timeout, name):
    inicialTime = time()

    while True:
        #try:
        # try to find image on screen
        location = pyautogui.locateOnScreen(image, confidence=0.9)

        if location is not None:
            # Image found
            logging.info(f'{name} image found.')
            return location

        #except:
        # checking if timeout is over
        else:
            currentTime = time()
            if currentTime - inicialTime >= timeout:
                # timeout is over and the image was not found.
                logging.info(f'{name} image not found.')
                return None

        # waiting for searching again
        sleep(0.5)
'''

def waitForImage(image, timeout, name):
    inicialTime = time()
    try:
        while True:
            # try to find image on screen
            location = pyautogui.locateOnScreen(image, confidence=0.9)

            if location is not None:
                # Image found
                logging.info(f'{name} image found.')
                return location

            #except:
            # checking if timeout is over
            else:
                currentTime = time()
                if currentTime - inicialTime >= timeout:
                    # timeout is over and the image was not found.
                    logging.info(f'{name} image not found.')
                    return None

            # waiting for searching again
            sleep(0.5)
    except:
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
                currentTime = time()
                if currentTime - inicialTime >= timeout:
                    # timeout is over and the image was not found.
                    logging.info(f'{name} image not found.')
                    return None

                # waiting for searching again
                sleep(0.1)


def emailSuccess(nome_bot, texto_msg):
    with open(f'{email_path}/emailSuccess.html', 'r', encoding='utf=8') as fileSuccess:
        templateSuccess = fileSuccess.read()
        htmlSuccess = templateSuccess.format(automacao = nome_bot, mensagem = texto_msg)
    yag.send(to=str(email_receivers).split(','), subject=f"TESTES - BOT FINALIZADO - RETORNO PAGAMENTO EM MOEDA ESTRANGEIRA - BV",
    contents=htmlSuccess,
    #attachments=[f'{log_path}/{today}/{log_file_name}.log', writer, writer_erro])
    attachments=[f'{log_path}/{today}/{log_file_name}.log', writer])
    logging.info('Success e-mail sent.')

    sys.exit(0)


def emailErro(texto_msg):
    with open(f'{email_path}/emailError.html', 'r', encoding='utf=8') as fileSuccess:
        templateSuccess = fileSuccess.read()
        htmlSuccess = templateSuccess.format(mensagem = texto_msg)
    yag.send(to=str(email_receivers).split(','), subject=f"TESTES - ERRO - BOT RETORNO PAGAMENTO EM MOEDA ESTRANGEIRA - BV",
    contents=htmlSuccess,
    attachments=[f'{log_path}/{today}/{log_file_name}.log', writer])
    logging.info('Warning e-mail sent.')

    sys.exit(-1)




def headerExcel(nome_aba='PROCESSADOS'):
    global workbook, sheet

    logging.info('Criando EXCEL')

    workbook = openpyxl.Workbook()

    # Acesse a planilha ativa (por padrão, há uma planilha criada)
    sheet = workbook.active
    sheet.title = nome_aba

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


    logging.info('EXCEL iniciado')

def createRowExcel1(i, data, pi, inv, moeda, valor, nome, cod, erro, status):

    logging.info(f'Criando linha do EXCEL {pi} - {data} - {inv} - {moeda} - {valor} - {nome} - {cod} - {erro} - {status}')
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


def createRowExcel(i, data, pi, inv, moeda, valor, nome, cod, erro, status):

    sheet = workbook.active
    
    # Encontra a primeira linha vazia na coluna A
    row_index = 1
    while sheet[f'A{row_index}'].value is not None:
        row_index += 1
    
    # Escreve os dados na próxima linha vazia
    sheet[f'A{row_index}'] = data
    sheet[f'B{row_index}'] = pi
    sheet[f'C{row_index}'] = inv
    sheet[f'D{row_index}'] = moeda
    sheet[f'E{row_index}'] = valor
    sheet[f'F{row_index}'] = nome
    sheet[f'G{row_index}'] = cod
    sheet[f'H{row_index}'] = erro
    sheet[f'I{row_index}'] = status


def inicial():

    pidKillFinish()
    bot_text_art()
    headerExcel()


def getPid():
    logging.info('Dentro função GetPid\n')
    pid_list = []
    name_list = []
    for processName in processList:
        processes = psutil.process_iter(attrs=['pid', 'name'])

        for process in processes:
            pid = process.info['pid']
            name = process.info['name']
            if processName in name: 
                logging.info(f"Process {processName} / PID: {pid} found.")
                pid_list.append(pid)
                name_list.append(processName)
    return pid_list, name_list


def pidKillFinish():
    pid_list, name_list = getPid()

    for pid, name in zip(pid_list, name_list):
        os.kill(pid, signal.SIGTERM)
        logging.info(f'Process {name} / PID: {pid} killed.')





def excelErro(erro_invoice):
    logging.info('Adicionando aba de ERRO ao EXCEL existente')

    # Criar uma nova aba no workbook existente
    sheet = workbook.create_sheet(title="Baixa manual necessaria")

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

    for i, extrato in enumerate(erro_invoice):
        #logging.info(f'PROCESSO {i+1} A SER TRATADO: {extrato}')
        
        data = extrato["data"].replace("-", "")
        pi = extrato["pi"]
        inv = extrato["inv"]
        moeda = extrato["moeda"]
        valor = extrato["saldo"]
        nome = extrato["nome"]
        cod = extrato["cod"]

        erro = 'Fora do padrão'
        status = 'Não baixado'

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
    #workbook.save(f'{writer_erro}')



def excelProcessoCompleto(lista_processos, arquivo_extrato):

    logging.info('Adicionando aba de PROCESSOS EXTRAÍDOS ao EXCEL existente')

    # Criar uma nova aba no workbook existente
    sheet = workbook.create_sheet(title=f"TODOS EXTRAÍDOS")

    # Escreva dados nas células
    sheet['A1'] = 'DATA'
    sheet['B1'] = 'PI'
    sheet['C1'] = 'INV'
    sheet['D1'] = 'MOEDA'
    sheet['E1'] = 'VALOR'
    sheet['F1'] = 'NOME'
    sheet['G1'] = 'COD'

    logging.info('Aba de TODOS EXTRAÍDOS ao EXCEL existente')

    for i, extrato in enumerate(lista_processos):
        #logging.info(f'PROCESSO {i+1} A SER ESCRITO: {extrato}')
        
        data = extrato["data"].replace("-", "")
        pi = extrato["pi"]
        inv = extrato["inv"]
        moeda = extrato["moeda"]
        valor = extrato["montante"].replace("-", "")
        nome = extrato["nome"]
        cod = extrato["cod"]

        # Escrever dados na nova aba
        sheet[f'A{i+2}'] = data
        sheet[f'B{i+2}'] = pi
        sheet[f'C{i+2}'] = inv
        sheet[f'D{i+2}'] = moeda
        sheet[f'E{i+2}'] = valor
        sheet[f'F{i+2}'] = nome
        sheet[f'G{i+2}'] = cod

    #logging.info('Fechando Excel com Erro...')
    #workbook.save(f'{writer_erro}')
     


def verificar_baixado(extrato):
    global processos_baixados

    with open('tratados.json', 'r') as arquivo:
        # Carrega o conteúdo do arquivo JSON em um dicionário Python
        processos_baixados = json.load(arquivo)

    if extrato in processos_baixados:

        erro = 'Tratado em execução anterior do BOT'
        return erro


def salvar_baixado(extrato):
    
    processos_baixados.append(extrato)

    with open('tratados.json', 'w') as arquivo:
        json.dump(processos_baixados, arquivo, indent=4)

    
def mover_para_lidos(cam_arquivo):

    arquivo = os.path.basename(cam_arquivo)

    logging.info(f'MOVENDO ARQUIVO {arquivo}.')
    arquivo_novo = arquivo

    lidos_destino = f'{pasta_local}\\Lidos'
    if  os.path.exists(lidos_destino) == False:
        os.makedirs(lidos_destino)
        logging.info(f'Criando pasta {lidos_destino}')
    else:
        logging.info(f'Pasta "Lidos" já existe')


    teste_caminho_destino = os.path.join(lidos_destino, arquivo)
    contador = 1
    while os.path.exists(teste_caminho_destino):
    #while arquivo in lidos_destino:
        nome_arquivo_sem_extensao, extensao = os.path.splitext(arquivo)
        arquivo_novo = f"{nome_arquivo_sem_extensao}_{contador}{extensao}"
        teste_caminho_destino = os.path.join(lidos_destino, arquivo_novo)
        contador += 1
        logging.info(f'Novo nome: {arquivo_novo}')
        logging.info(f'teste_caminho_destino: {teste_caminho_destino}')
    
    caminho_destino = os.path.join(lidos_destino, arquivo_novo)

    shutil.move(cam_arquivo, caminho_destino)
    logging.info(f'ARQUIVO {arquivo_novo} MOVIDO PARA LIDOS.')