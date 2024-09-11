
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
import openpyxl
from openpyxl.styles import Font, Border, Fill, Alignment, Protection

from openpyxl.styles import PatternFill, Font

from copy import copy

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

def copiar_aba(origem, destino):
    for row in origem.iter_rows():
        for cell in row:
            new_cell = destino.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)

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
            wb_final = openpyxl.Workbook()
            wb_final.remove(wb_final.active)  # Remove a aba padrão
            abas = [caminho_arquivo_b, caminho_arquivo_c, caminho_arquivo_d]
            nomes_abas = ['Apuração do Faturamento', 'Gabarito Prateleira - SIAD', 'Gabarito Órgão - Sigla 2024']

            for arquivo, nome_aba in zip(abas, nomes_abas):
                wb_origem = openpyxl.load_workbook(arquivo, data_only=True)
                ws_origem = wb_origem.active
                ws_destino = wb_final.create_sheet(title=nome_aba)
                copiar_aba(ws_origem, ws_destino)

            nome_arquivo_final = os.path.join(caminho_pasta, f"Valor Faturado - 20{mes_ano}.xlsx")
            wb_final.save(nome_arquivo_final)
            print(f"Arquivo salvo como: {nome_arquivo_final}")
        else:
            print("Um ou mais arquivos não foram encontrados.")
    else:
        print("Necessário fornecer o nome da pasta e a data.")

if __name__ == "__main__":
    main()
