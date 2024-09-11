# a fazer:
##
##
##
##

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


def main():
    nomepasta = entrada_nomepasta.get()
    caminhoDesktop = encontrar_caminho_area_de_trabalho()

    if not nomepasta:
        messagebox.showerror("Erro", "Preencha todas as caixas de texto.")
        return
    


    encontrar_caminho_area_de_trabalho()
    df_apuracao = encontrar_caminho_apuracao()
    df_gabarito = encontrar_caminho_cod_orgao()
    df_cod_orgao = encontrar_caminho_gabarito()

    if df_apuracao is None or df_gabarito is None or df_cod_orgao is None:
        messagebox.showerror("Erro", "Um ou mais arquivos corretos não foram encontrados na pasta.")
        return

    print("Todos os arquivos foram lidos com sucesso.")
    # lembrar de fazer um os.path join com o caminho do arquivo novo


def encontrar_caminho_area_de_trabalho():
    caminhos_possiveis = [
        os.path.join(os.path.expanduser("~"),"Desktop"),
        os.path.join(os.path.expanduser("~"),"Área de Trabalho"),
        os.path.join(os.path.expanduser("~"),"OneDrive", "Área de Trabalho"),
        os.path.join(os.path.expanduser("~"),"OneDrive", "Desktop")
    ]
    for caminho in caminhos_possiveis:
        if os.path.exists(caminho):
            return caminho
    raise FileExistsError("Não foi possível encontrar a pasta Área de Trabalho ou Desktop.") #O raise está fora do loop for e não diretamente dentro do if porque o objetivo é verificar todos os caminhos possíveis antes de decidir se uma exceção deve ser lançada.



def encontrar_caminho_apuracao():
    caminhoDesktop = encontrar_caminho_area_de_trabalho()
    nomepasta = entrada_nomepasta.get()
    nomepasta = nomepasta.strip()
    caminho_simples = os.path.join(caminhoDesktop, nomepasta) #nome vai ser fornecido via tk
    nome_apuracao = [f for f in os.listdir(caminho_simples) if f.startswith('Apuração') and f.endswith('xlsx')]

    if not nome_apuracao:
            print("Nenhum arquivo de 'Apuração do faturamento' encontrado.")
            return None  # Retorna None se não encontrar arquivos

    try:
        for arquivo in nome_apuracao:
            caminho_apuracao = os.path.join(caminho_simples, arquivo)
            df_apuracao = pd.read_excel(caminho_apuracao, sheet_name = 'Apuração do Faturamento')
            print(f"Arquivo {arquivo} lido com sucesso")
        return df_apuracao
    except PermissionError:
        messagebox.showerror("Erro", "Erro ao ler os arquivos. Verifique se o arquivo Apuração do Faturamento não esteja aberto")
        return

    



def encontrar_caminho_gabarito():
    caminhoDesktop = encontrar_caminho_area_de_trabalho()
    nomepasta = entrada_nomepasta.get()
    nomepasta = nomepasta.strip()
    caminho_simples = os.path.join(caminhoDesktop, nomepasta)
    nome_gabarito = [f for f in os.listdir(caminho_simples) if f.startswith('Gabarito Prateleira') and f.endswith('xlsx')]

    if not nome_gabarito:
        print("Nenhum arquivo de 'Gabarito Prateleira' encontrado.")
        return None

    try:
        for arquivo in nome_gabarito:
            caminho_gabarito = os.path.join(caminho_simples, arquivo)
            df_gabarito = pd.read_excel(caminho_gabarito)
            print(f"Arquivo {arquivo} lido com sucesso")
        return df_gabarito
    except PermissionError:
        messagebox.showerror("Erro", "Erro ao ler os arquivos. Verifique se o arquivo gabarito taltal não esteja aberto")
        return



def encontrar_caminho_cod_orgao():
    caminhoDesktop = encontrar_caminho_area_de_trabalho()
    nomepasta = entrada_nomepasta.get()
    nomepasta = nomepasta.strip()
    caminho_simples = os.path.join(caminhoDesktop, nomepasta)
    nome_cod_orgao = [f for f in os.listdir(caminho_simples) if f.startswith('Gabarito Órgão-Sigla') and f.endswith('xlsx')]

    if not nome_cod_orgao:
        print("Nenhum arquivo de 'Gabarito Órgão-Sigla' encontrado.")
        return None

    try:
        for arquivo in nome_cod_orgao:
            caminho_cod_orgao = os.path.join(caminho_simples, arquivo)
            df_cod_orgao = pd.read_excel(caminho_cod_orgao)
            print(f"Arquivo {arquivo} lido com sucesso")
        return df_cod_orgao
    except PermissionError:
        messagebox.showerror("Erro", "Erro ao ler os arquivos. Verifique se o arquivo Código taltal não esteja aberto")
        return



# A partir daqui o código é referente à janela 

janela = tk.Tk()
janela.geometry('400x200')
janela.title('Valor Faturado') #rever nome da janela
janela.resizable(False,False) #pra não conseguirem mudar o tamanho da caixa
texto = tk.Label(janela, text='Nome pasta:')
texto.grid(column=0, row=0, pady=10) #tentar padx dps pra ver a diferença
entrada_nomepasta = tk.Entry(janela)
entrada_nomepasta.grid(column=1, row=0, ipadx = 60)
#ainda tenho que colocar uma entrada para o nome do arquivo
botao = tk.Button(janela, text='Processar', command=main)
##botao.grid(column = 0, row=3, pady=10)
botao.grid(column=0, row=3, columnspan=2, pady=20, ipadx=20, ipady=5)


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