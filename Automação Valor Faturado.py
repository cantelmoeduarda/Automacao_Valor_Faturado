''' 
A fazer:
- colocar opção de data na janela (usando caixinha de data)
- Fechar janela depois de clicar em aplicar
- Juntar a alteracao_planilha no código principal
- Fazer o design da janela

'''

#importações
import os
from datetime import datetime as dt
from dateutil.relativedelta import relativedelta
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
from tkinterdnd2 import DND_FILES, TkinterDnD
import pyautogui as bot
from openpyxl.styles import NamedStyle, Border, Side
from openpyxl.utils import get_column_letter
import sys
import pandas as pd
from tkinter import simpledialog
from tkinter.filedialog import askopenfile
from tkinter.filedialog import askdirectory



def main():


    df_apuracao = encontrar_caminho_apuracao()
    df_prateleira = encontrar_caminho_cod_orgao()
    df_orgao_sigla = encontrar_caminho_gabarito()
    

    if df_apuracao is None or df_prateleira is None or df_orgao_sigla is None:
        messagebox.showerror("Erro", "Um ou mais arquivos corretos não foram encontrados na pasta.")
        raise Exception('Um dos arquivos não foram encontrados') #isso aqui vai parar o código
    
    print("Todos os arquivos foram lidoss com sucesso.")

    
    return


def encontrar_caminho_apuracao():
    caminho_pasta = var_caminho_pasta.get()
    nome_apuracao = [f for f in os.listdir(caminho_pasta) if f.startswith('Apuração') and f.endswith('xlsx')]

    if not nome_apuracao:
            print("Nenhum arquivo de 'Apuração do faturamento' encontrado.")
            return None  # Retorna None se não encontrar arquivos

    try:
        for arquivo in nome_apuracao:
            caminho_apuracao = os.path.join(caminho_pasta, arquivo)
            df_apuracao = pd.read_excel(caminho_apuracao, sheet_name = 'Apuração do Faturamento')
            print(f"Arquivo {arquivo} lido com sucesso")
        return df_apuracao
    except PermissionError:
        messagebox.showerror("Erro", "Erro ao ler os arquivos. Verifique se o arquivo Apuração do Faturamento não esteja aberto")
        raise Exception('Erro ao ler arquivo Apuração do Faturamento')

    



def encontrar_caminho_gabarito():
    caminho_pasta = var_caminho_pasta.get()
    nome_gabarito = [f for f in os.listdir(caminho_pasta) if f.startswith('Gabarito Prateleira') and f.endswith('xlsx')]

    if not nome_gabarito:
        print("Nenhum arquivo de 'Gabarito Prateleira' encontrado.")
        return None

    try:
        for arquivo in nome_gabarito:
            caminho_gabarito = os.path.join(caminho_pasta, arquivo)
            df_prateleira = pd.read_excel(caminho_gabarito)
            print(f"Arquivo {arquivo} lido com sucesso")
        return df_prateleira
    except PermissionError:
        messagebox.showerror("Erro", "Erro ao ler os arquivos. Verifique se o arquivo gabarito taltal não esteja aberto")
        raise Exception('Erro ao ler arquivo gabarito taltal')



def encontrar_caminho_cod_orgao():
    caminho_pasta = var_caminho_pasta.get()
    nome_cod_orgao = [f for f in os.listdir(caminho_pasta) if f.startswith('Gabarito Órgão-Sigla') and f.endswith('xlsx')]

    if not nome_cod_orgao:
        print("Nenhum arquivo de 'Gabarito Órgão-Sigla' encontrado.")
        return None

    try:
        for arquivo in nome_cod_orgao:
            caminho_cod_orgao = os.path.join(caminho_pasta, arquivo)
            df_orgao_sigla = pd.read_excel(caminho_cod_orgao)
            print(f"Arquivo {arquivo} lido com sucesso")
        return df_orgao_sigla
    except PermissionError:
        messagebox.showerror("Erro", "Erro ao ler os arquivos. Verifique se o arquivo Código taltal não esteja aberto")
        raise Exception('Erro ao ler arquivo código taltal')



# A partir daqui o código é referente à janela 

def selecionar_arquivo():
    caminho_pasta = askdirectory(title='Selecione a pasta com os arquivos')
    var_caminho_pasta.set(caminho_pasta)    
    if caminho_pasta:
        label_pasta_selecionada['text'] = f"Pasta selecionada: {os.path.basename(caminho_pasta)}" #ve se isso rola

janela = tk.Tk()
janela.geometry('400x200')
janela.title('Valor Faturado') #rever nome da janela
janela.resizable(False,False) #pra não conseguirem mudar o tamanho da caixa
texto = tk.Label(janela, text='Selecione a pasta com os arquivos:')
texto.grid(column=0, row=0, pady=10) 

var_caminho_pasta = tk.StringVar()
botao_selecionararquivo = tk.Button(janela, text="Clique para Selecionar", command=selecionar_arquivo)
botao_selecionararquivo.grid(row=3, column=0, padx=10, pady=10, sticky='nsew')
label_pasta_selecionada = tk.Label(janela, text='Nenhuma pasta selecionada', anchor='e')
label_pasta_selecionada.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky='nsew')


botao_processar = tk.Button(janela, text='Processar', command=main)
botao_processar.grid(column=0, row=5, columnspan=2,ipady=5)


# Tratando as imagens que farão parte do botão:
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path,relative_path)

# icone triagulo
image_path = resource_path("triangulo.png")
imagem = Image.open(image_path)

imagem = imagem.resize((32,32),Image.LANCZOS)
icone = ImageTk.PhotoImage(imagem)
janela.iconphoto(True, icone)

# botao de ajuda
def ajuda():
    messagebox.showinfo("Informações importantes",
                        "Selecione taltaltal... (usar barra + n para colocar texto embaixo)")
    
icon_path = resource_path("botao_de_ajuda_transparente.png")
help_icon = Image.open(icon_path)
help_icon = help_icon.resize((22, 22), Image.LANCZOS)
help_icon = ImageTk.PhotoImage(help_icon)
help_button = tk.Button(janela, image=help_icon, command=ajuda, borderwidth=0, bg='#f0f0f0', activebackground='#f0f0f0')
help_button.grid(column=1, row=3, sticky='e', padx=10)



janela.mainloop()
