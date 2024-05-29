from SUBPROGRAMS.parameters import *


def mover_para_lidos(cam_arquivo):

    arquivo = os.path.basename(cam_arquivo)

    logging.info(f'MOVENDO ARQUIVO {arquivo}.')
    arquivo_novo = arquivo

    lidos_destino = f'{teste_path}\\Lidos'
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
    




teste_path = f'{base_path}\\TESTE'

arquivos_pdf_bruto = glob.glob(os.path.join(teste_path, '*.pdf'))
logging.info(f'DDDDD {arquivos_pdf_bruto} ARQUIVOS.')



arquivos_pdf = []
for arquivo_bruto in arquivos_pdf_bruto:
    nom_arquivo = os.path.basename(arquivo_bruto)
    if 'parcial' in nom_arquivo.lower():
        arquivos_pdf.append(arquivo_bruto)
    
logging.info(f'SERÃO TRATADOS {len(arquivos_pdf)} ARQUIVOS.')
logging.info(f'ARQUIVOS: {arquivos_pdf}')



for cam_arquivo in arquivos_pdf:
    mover_para_lidos(cam_arquivo)

