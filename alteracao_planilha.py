#alteração das planilhas sem o botão
import pandas as pd
df_apuracao = pd.read_excel('C:/Users/X15848318638/Desktop/primeira/Apuração do faturamento 202308.xlsx', sheet_name = 'Apuração do Faturamento')
df_gabarito = pd.read_excel('C:/Users/X15848318638/Desktop/primeira/Gabarito Órgão-Sigla 202406.xlsx')
df_cod_orgao = pd.read_excel('C:/Users/X15848318638/Desktop/primeira/Gabarito Prateleira - SIAD - 202409.xlsx')

posicao_col_valor_faturar = df_apuracao.columns.get_loc('Valor a Faturar pós IMR')
#print(f'A posição é {posicao_col_valor_faturar}')

df_apuracao.insert(58,'Ajustes',0)
#df_apuracao.info()
print(df_apuracao)