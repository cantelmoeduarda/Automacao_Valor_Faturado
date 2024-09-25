
#importações
import os
from datetime import datetime as dt
from dateutil.relativedelta import relativedelta
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
from tkinterdnd2 import DND_FILES, TkinterDnD
import pyautogui as bot
import openpyxl as xl
from openpyxl.styles import NamedStyle, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
import sys
import pandas as pd
from tkinter import simpledialog
from tkinter.filedialog import askopenfile
from tkinter.filedialog import askdirectory



def main():

    caminho_pasta = var_caminho_pasta.get() 
    df_apuracao = encontrar_caminho_apuracao()
    df_orgao_sigla = encontrar_caminho_orgao_sigla()
    df_prateleira = encontrar_caminho_gabarito()
    

    if df_apuracao is None or df_prateleira is None or df_orgao_sigla is None:
        messagebox.showerror("Erro", "Um ou mais arquivos corretos não foram encontrados na pasta.")
        raise Exception('Um dos arquivos não foram encontrados') #isso aqui vai parar o código
    
    
    messagebox.showinfo('Sucesso',f'Arquivo Valor Faturado salvo na pasta "{os.path.basename(caminho_pasta)}" com sucesso')
    df_apuracao.insert(58,'Ajustes',0)
    # Etapa de preenchimento de uma nova coluna SIGLA com valores da tabela gabarito órgão sigla
    df_orgao_sigla = df_orgao_sigla.rename(columns = {'Cliente Nome':'Órgão/Entidade'})
    df_apuracao = df_apuracao.merge(df_orgao_sigla[['Órgão/Entidade','Sigla']],on='Órgão/Entidade', how='left') #esse merge deu tudo certo


    # Etapa de preenchimento de uma nova coluna Item e Categoria com valores da tabela gabarito prateleira
    df_prateleira.drop(['Código Item Material - Numérico','Item Material','Item Correspondente.1','Situação'],axis='columns',inplace=True)
    df_prateleira.drop_duplicates(subset=['Código do item AVMG'],inplace=True)
    df_prateleira = df_prateleira.rename(columns = {'Código do item AVMG':'ID Item'})
    df_prateleira = df_prateleira.rename(columns = {'Item Correspondente':'Item'})
    df_apuracao = df_apuracao.merge(df_prateleira,on='ID Item', how='left')


    # Etapa Preencheendo  a coluna “Observação” com o texto “Sem observação”
    df_apuracao['Observação'] = 'Sem observação'


    # Etapa Preencheendo  a coluna “Exercício” com o ano referente ao da coluna Data da Aprovação
    df_apuracao['Exercício'] = df_apuracao['Data da Aprovação'].dt.year.fillna(0).astype(int)


    # Etapa de exclusão de valores "Não se aplica"
    df_apuracao['Data limite de entrega - pedido original'] = df_apuracao['Data limite de entrega - pedido original'].replace('Não se aplica', "")
    df_apuracao['Dias de atraso - pedido original'] = df_apuracao['Dias de atraso - pedido original'].replace('Não se aplica', "")
    df_apuracao['Data limite de entrega - entrega corretiva'] = df_apuracao['Data limite de entrega - entrega corretiva'].replace('Não se aplica', "")
    df_apuracao['Dias de atraso - entrega corretiva'] = df_apuracao['Dias de atraso - entrega corretiva'].replace('Não se aplica', "")


    
    #Salvando o arquivo em excel com algumas colunas em formato de data e design em geral
    caminho_arquivo_final = os.path.join(caminho_pasta,'Valor Faturado.xlsx')
    with pd.ExcelWriter(caminho_arquivo_final, engine='openpyxl') as writer:
        df_apuracao.to_excel(writer, index=False, sheet_name='Valor Faturado')

    wb = load_workbook(caminho_arquivo_final)
    ws = wb['Valor Faturado']
    
    #Etapa de tirar o time da data
    colunas_de_data = ['Data do Fato Gerador', 'Data da Aprovação', 'Data do Ateste','Data limite de entrega - entrega corretiva', 'Data limite de entrega - pedido original']
    for data in colunas_de_data:
        df_apuracao[data] = pd.to_datetime(df_apuracao[data]).dt.strftime('%d/%m/%Y')

    if 'date_style' not in wb.named_styles:
        date_style = NamedStyle(name='date_style', number_format='DD/MM/YYYY')
        wb.add_named_style(date_style)

    colunas_de_data_excel = ['BJ', 'AV', 'AR', 'AQ']
    for col in colunas_de_data:
        if col in df_apuracao.columns:
            df_apuracao[col] = pd.to_datetime(df_apuracao[col], format='%d/%m/%Y', errors='coerce')
            
    for col in colunas_de_data_excel:
        for row in range(2, ws.max_row + 1): 
            cell = ws[f'{col}{row}']
            if cell.value:
                cell.style = 'date_style'
            
            
    filtro = f"A1:{get_column_letter(ws.max_column)}1"
    ws.auto_filter.ref = filtro

    wb.save(caminho_arquivo_final)
    
    
    
    return


def encontrar_caminho_apuracao():
    caminho_pasta = var_caminho_pasta.get()
    nome_apuracao = [f for f in os.listdir(caminho_pasta) if (f.startswith('Apuração')
                     or f.startswith('Apuração Faturamento')
                     or f.startswith('Apuracao do Faturamento')
                     or f.startswith('Apuracao faturamento')
                     or f.startswith('Apuração do faturamento'))
                     and f.endswith('xlsx')]

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
    nome_gabarito = [f for f in os.listdir(caminho_pasta) if (f.startswith('Gabarito Prateleira')
                     or f.startswith('gabarito prateleira')
                     or f.startswith('Gabarito-Prateleira')
                     or f.startswith('gabarito-prateleira'))
                     and f.endswith('xlsx')]

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
        messagebox.showerror("Erro", "Erro ao ler os arquivos. Verifique se o arquivo Gabarito Prateleira não esteja aberto")
        raise Exception('Erro ao ler arquivo gabarito Gabarito Prateleira')



def encontrar_caminho_orgao_sigla():
    caminho_pasta = var_caminho_pasta.get()
    nome_orgao_sigla = [f for f in os.listdir(caminho_pasta) if (f.startswith('Gabarito Órgão-Sigla') 
                        or f.startswith('Gabarito Órgão Sigla')
                        or f.startswith('Gabarito órgão-sigla')
                        or f.startswith('Gabarito Orgão-Sigla'))
                        and f.endswith('xlsx')]

    if not nome_orgao_sigla:
        print("Nenhum arquivo de 'Gabarito Órgão-Sigla' encontrado.")
        return None

    try:
        for arquivo in nome_orgao_sigla:
            caminho_orgao_sigla = os.path.join(caminho_pasta, arquivo)
            df_orgao_sigla = pd.read_excel(caminho_orgao_sigla)
            print(f"Arquivo {arquivo} lido com sucesso")
        return df_orgao_sigla
    except PermissionError:
        messagebox.showerror("Erro", "Erro ao ler os arquivos. Verifique se o arquivo Gabarito Órgão-Sigla esteja aberto")
        raise Exception('Erro ao ler Gabarito Órgão-Sigla')



# A partir daqui o código é referente à janela 

#Essa função 
def selecionar_arquivo():
    caminho_pasta = askdirectory(title='Selecione a pasta com os arquivos')
    var_caminho_pasta.set(caminho_pasta)    
    if caminho_pasta:
        label_pasta_selecionada['text'] = f"* Pasta selecionada: {os.path.basename(caminho_pasta)}" #ve se isso rola
    

janela = tk.Tk()
janela.geometry('400x200')
janela.title('Valor Faturado') #rever nome da janela
janela.resizable(False,False) #pra não conseguirem mudar o tamanho da caixa
texto = tk.Label(janela, text='Selecione a pasta com os arquivos:')
texto.grid(column=1, row=0, padx=10, pady=10, sticky='w')

janela.grid_columnconfigure(0, weight=1)
janela.grid_columnconfigure(1, weight=1)

var_caminho_pasta = tk.StringVar()
botao_selecionararquivo = tk.Button(janela, text="Selecionar", command=selecionar_arquivo)
botao_selecionararquivo.grid(row=1, column=1, padx=10, pady=10, sticky='nsew')
label_pasta_selecionada = tk.Label(janela, text='* Nenhuma pasta selecionada',fg='blue')
label_pasta_selecionada.grid(row=2, column=1, sticky='nsew')

botao_processar = tk.Button(janela, text='Processar', command=main)
botao_processar.grid(column=1, row=3, columnspan=1, padx=10, pady=10, ipady=5, sticky='nsew')



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
                        "Selecione a pasta que contenha os arquivos:\n\nApuração do Faturamento\nGabarito Prateleira - SIAD\nGabarito Órgão - Sigla")
    
icon_path = resource_path("botao_de_ajuda_transparente.png")
help_icon = Image.open(icon_path)
help_icon = help_icon.resize((22, 22), Image.LANCZOS)
help_icon = ImageTk.PhotoImage(help_icon)
help_button = tk.Button(janela, image=help_icon, command=ajuda, borderwidth=0, bg='#f0f0f0', activebackground='#f0f0f0')
help_button.grid(row=4, column=4, padx=10, sticky='e')


janela.mainloop()
