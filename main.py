'''from clicknium import locator, ui
from clicknium import clicknium as cc
import json'''

from SUBPROGRAMS.functions import *


def process():

    global arquivo

    inicial()

    arquivos_pdf_bruto = glob.glob(os.path.join(pasta_local, '*.pdf'))
    logging.info(f'TODOS ARQUIVOS NA PASTA {arquivos_pdf_bruto}.')
    
    arquivos_pdf = []
    for arquivo_bruto in arquivos_pdf_bruto:
        nom_arquivo = os.path.basename(arquivo_bruto)
        if 'parcial' in nom_arquivo.lower():
            arquivos_pdf.append(arquivo_bruto)
        
    logging.info(f'SERÃO TRATADOS {len(arquivos_pdf)} ARQUIVOS.')
    logging.info(f'ARQUIVOS: {arquivos_pdf}')

    if not arquivos_pdf:
        texto_msg = f'NENHUM PDF "PARCIAL" DISPONÍVEL'
        logging.info(texto_msg)
        encerra_prog()
        emailErro(texto_msg)

    for cont, arquivo in enumerate(arquivos_pdf):
        logging.info(f'INICIO DO PROCESSAMENTO DO ARQUIVO {cont+1}: {arquivo}')

        #erro_invoice = ''
        #extrato_list_nao_tratado = informacoes_extraidas()
        #extrato_list = json.loads(extrato_list_nao_tratado)
        
        extrato_list_nao_tratado, erro_invoice = informacoes_extraidas(arquivo)


        logging.info(f'PROCESSOS EXTRAIDOS COM ERRO: {len(erro_invoice)}')
        logging.info(f'PROCESSOS COM ERRO: {erro_invoice}')
        excelErro(erro_invoice)

        logging.info(f'PROCESSOS EXTRAIDOS CORRETAMENTE: {len(extrato_list_nao_tratado)}')
        logging.info(f'PROCESSOS A SEREM TRATADOS: {extrato_list_nao_tratado}')

        #logging.info(f'PROCESSOS EXTRAIDOS CORRETAMENTE: {len(extrato_list)}')
        #logging.info(f'PROCESSOS A SEREM TRATADOS: {extrato_list}')

        mover_para_lidos(arquivo) 

        if cont == 0: 
            open_RM()
            loggin_RM()
            enter_coligada()
            enter_contas_a_pagar()
        
        flag = False
        for i, extrato in enumerate(extrato_list_nao_tratado):
        #for i, extrato in enumerate(extrato_list):
            
            logging.info(f'extrato_list_nao_tratado processo {i}')
            

            if i != 0:
                logging.info(f'PASSANDO PARA PROXIMO PROCESSO')


            logging.info(f'PROCESSO {i+1} de {len(extrato_list_nao_tratado)} A SER TRATADO: {extrato}')

            

            data = extrato["data"]
            for _ in range(2):
                data = data.replace("-", "")

            pi = extrato["pi"]

            inv_extraida = extrato["inv"]
            logging.info(f'inv original {inv_extraida}. tamanho {len(inv_extraida)}')
            
            inv = f"{inv_extraida}01"
            while len(inv) < 11:
                inv = f"0{inv}"
            logging.info(f'inv TRATADA {inv}. tamanho {len(inv)}')

            moeda = extrato["moeda"]

            valor = extrato["montante"]
            valor = valor.replace("-", "")
            valor = valor.replace(".", ",")

            saldo = extrato["saldo"]
            nome = extrato["nome"]
            cod = extrato["cod"]

            logging.info(f"""data: {data} {type(data)}\n
                        pi: {pi} {type(pi)}\n
                        inv: {inv} {type(inv)}\n
                        moeda: {moeda} {type(moeda)}\n
                        montante: {valor} {type(valor)}\n
                        saldo: {saldo} {type(saldo)}\n
                        nome: {nome} {type(nome)}\n
                        cod: {cod} {type(cod)}""")
            

            erro = verificar_baixado(extrato)
            if erro != "" and erro != None:
                logging.info(f'{erro}')
                status = 'Baixado anteriormente'
                createRowExcel(i, data, pi, inv_extraida, moeda, valor, nome, cod, erro, status) 

                if i == 0 and cont == 0:
                    flag = True

                continue
            
            
            filter_lancamento(inv, i, cont, flag)
            if i == 0 and cont == 0 or flag:
                desmark_filter()
                filter_clean()
                select_tipo_documento()
                flag = False


            result = check_table_record()
            if result == True:
                
                erro = 'Nenhuma INV encontrada'
                status = 'Não Baixado'
                logging.info(erro)

                createRowExcel(i, data, pi, inv_extraida, moeda, valor, nome, cod, erro, status)             
                continue
            

            erro = check_inv(inv, nome, valor)
            if erro != "" and erro != None:
                logging.info(f'{erro}')
                status = 'Não Baixado'
                createRowExcel(i, data, pi, inv_extraida, moeda, valor, nome, cod, erro, status)  
                continue

            define_moeda(moeda)
            
            erro = baixa_inv(valor, moeda, data)
            #### CONFERIR POSSIVEL DATA E VALORES DIVERGENTES E ADEQUAR AO ERRO
            if erro != "" and erro != None:
                logging.info(f'{erro}')
                status = 'Não Baixado'
                createRowExcel(i, data, pi, inv_extraida, moeda, valor, nome, cod, erro, status)  
                continue
            
            erro = ""
            status = 'Baixado'
            createRowExcel(i, data, pi, inv_extraida, moeda, valor, nome, cod, erro, status)  
            logging.info(f'{status}')

            salvar_baixado(extrato)
        
        logging.info(f'ARQUIVO FINALIZADO: {arquivo}')
        
    

                


if __name__ == '__main__':

    sys.excepthook = show_exception_and_exit
    process()
    encerra_prog()
    emailSuccess(nome_bot, texto_msg)

