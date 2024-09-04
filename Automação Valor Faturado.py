import openpyxl as xl
from openpyxl.styles import PatternFill, Font
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
import tkinter as tk
from tkinter import simpledialog
import os

import pandas as pd
import tkinter as tk
from tkinter import simpledialog
import os

def encontrar_caminho_area_de_trabalho():
    caminhos_possiveis = [
        os.path.join(os.path.expanduser("~"), "Desktop"),
        os.path.join(os.path.expanduser("~"), "Área de Trabalho"),
        os.path.join(os.path.expanduser("~"), "OneDrive", "Área de Trabalho"),
        os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop"),
    ]

    for caminho in caminhos_possiveis:
        if os.path.exists(caminho):
            return caminho
    raise FileNotFoundError("Não foi possível encontrar a pasta Área de Trabalho ou Desktop.")

def encontrar_arquivo(pasta, nome_arquivo):
    caminho_desktop = encontrar_caminho_area_de_trabalho()
    caminho_completo = os.path.join(caminho_desktop, pasta, nome_arquivo)
    if os.path.exists(caminho_completo):
        return caminho_completo
    else:
        return None

def ler_arquivo(caminho_arquivo):
    try:
        df = pd.read_excel(caminho_arquivo)
        return df
    except Exception as e:
        return f"Erro ao ler o arquivo {caminho_arquivo}: {e}"

def main():
    root = tk.Tk()
    root.withdraw()

    pasta = simpledialog.askstring("Pasta", "Digite o nome da pasta na área de trabalho:", parent=root)
    mes_ano = simpledialog.askstring("Data", "Digite o mês e ano no formato MMYYYY (ex: 072024):", parent=root)

    if pasta and mes_ano:
        caminho_desktop = encontrar_caminho_area_de_trabalho()
        caminho_pasta = os.path.join(caminho_desktop, pasta)
        caminho_arquivo_b = encontrar_arquivo(pasta, 'B.xlsx')
        caminho_arquivo_c = encontrar_arquivo(pasta, 'C.xlsx')
        caminho_arquivo_d = encontrar_arquivo(pasta, 'D.xlsx')

        if caminho_arquivo_b and caminho_arquivo_c and caminho_arquivo_d:
            df_b = ler_arquivo(caminho_arquivo_b)
            df_c = ler_arquivo(caminho_arquivo_c)
            df_d = ler_arquivo(caminho_arquivo_d)

            if isinstance(df_b, pd.DataFrame) and isinstance(df_c, pd.DataFrame) and isinstance(df_d, pd.DataFrame):
                nome_arquivo_final = os.path.join(caminho_pasta, f"Valor Faturado - {mes_ano}.xlsx")
                with pd.ExcelWriter(nome_arquivo_final) as writer:
                    df_b.to_excel(writer, sheet_name='Apuração do Faturamento', index=False)
                    df_c.to_excel(writer, sheet_name='Gabarito Prateleira - SIAD', index=False)
                    df_d.to_excel(writer, sheet_name='Gabarito Órgão - Sigla 2024', index=False)
                print(f"Arquivo salvo como: {nome_arquivo_final}")
            else:
                print("Erro ao processar os DataFrames.")
        else:
            print("Um ou mais arquivos não foram encontrados.")
    else:
        print("Necessário fornecer o nome da pasta e a data.")

if __name__ == "__main__":
    main()
