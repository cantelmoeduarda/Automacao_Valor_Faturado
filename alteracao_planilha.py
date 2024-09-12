#alteração das planilhas sem o botão
import pandas as pd
df_apuracao = pd.read_excel('C:/Users/cante/OneDrive/Desktop/primeira/Apuração do faturamento 202408.xlsx', sheet_name = 'Apuração do Faturamento')
df_gabarito = pd.read_excel('C:/Users/cante/OneDrive/Desktop/primeira/Gabarito Órgão-Sigla 202409.xlsx')
df_cod_orgao = pd.read_excel('C:/Users/cante/OneDrive/Desktop/primeira/Gabarito Prateleira - SIAD - 202409.xlsx')

#posicao_col_valor_faturar = df_apuracao.columns.get_loc('Valor a Faturar pós IMR')
#print(f'A posição é {posicao_col_valor_faturar}')

# Etapa de criação da coluna Ajustes e 0 (zero) em toda coluna
df_apuracao.insert(58,'Ajustes',0)
# Etapa de preenchimento de uma nova coluna SIGLA com valores da tabela gabarito órgão sigla
df_gabarito = df_gabarito.rename(columns = {'Cliente Nome':'Órgão/Entidade'})
df_apuracao = df_apuracao.merge(df_gabarito[['Órgão/Entidade','Sigla']],on='Órgão/Entidade', how='left')
# Etapa de preenchimento de uma nova coluna Item com valores da tabela gabarito prateleira
df_cod_orgao = df_cod_orgao.rename(columns = {'Código do item AVMG':'ID Item'})
df_cod_orgao = df_cod_orgao.rename(columns = {'Item Correspondente':'Item'})
df_apuracao = df_apuracao.merge(df_cod_orgao[['ID Item','Item']],on='ID Item', how='left')
# Etapa de preenchimento de uma nova coluna Categoria com valores da tabela gabarito prateleira
df_apuracao = df_apuracao.merge(df_cod_orgao[['ID Item','Categoria']],on='ID Item', how='left')
# Etapa Preencheendo  a coluna “Observação” com o texto “Sem observação”
df_apuracao['Observação'] = 0
#print(df_apuracao)
df_apuracao.info()
print(df_apuracao)




#df_apuracao.info()
#print(df_apuracao)