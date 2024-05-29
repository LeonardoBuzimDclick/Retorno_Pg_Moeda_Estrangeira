import tkinter as tk
from tkinter import ttk
import re
from tkinter import messagebox
import os




def finalizar():
    global item_selecionado

    item_selecionado = combo.get()

    if item_selecionado == "":
        messagebox.showerror("MÊS NÃO SELECIONADO.")

    else:
        print(f"Item selecionado: {item_selecionado}")
        janela_principal.destroy()



def janela():
    global combo, janela_principal, itens

    janela_principal = tk.Tk()

    rotulo = tk.Label(janela_principal, text="""RETORNO BV
Mês do próximo processamento""")
    rotulo.pack(pady=30)

    largura = 200
    altura = 250

    largura_tela = janela_principal.winfo_screenwidth()
    altura_tela = janela_principal.winfo_screenheight()

    pos_x = (largura_tela - largura) // 2
    pos_y = (altura_tela - altura) // 2

    janela_principal.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

    rotulo = tk.Label(janela_principal, text="Selecione o mês:")
    rotulo.pack(pady=0)

    itens = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO',
                         'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 
                                'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO']
    
    combo = ttk.Combobox(janela_principal, values=itens, state="readonly")
    combo.pack(pady=5)

    botao_mostrar = tk.Button(janela_principal, text="Salvar", command=finalizar)
    botao_mostrar.pack(pady=5)

    # Iniciar o loop principal da interface gráfica
    janela_principal.mainloop()

def escrever_arquivo():

    mes = str((itens.index(item_selecionado))+1)
    print(f"Mes: {mes}")

    with open('C:\Projetos_Retornos\Retorno_Pg_Moeda_Estrangeira\mes.txt', 'w') as arquivo:
        
        
        # Escreve a informação fornecida no início do arquivo
        arquivo.write(mes)

janela()
escrever_arquivo()




