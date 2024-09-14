#alteração das planilhas sem o botão
import pandas as pd
import openpyxl
#do pessoal:
df_apuracao = pd.read_excel('C:/Users/cante/OneDrive/Desktop/primeira/Apuração do faturamento 202408.xlsx', sheet_name = 'Apuração do Faturamento')
df_gabarito = pd.read_excel('C:/Users/cante/OneDrive/Desktop/primeira/Gabarito Órgão-Sigla 202409.xlsx')
df_cod_orgao = pd.read_excel('C:/Users/cante/OneDrive/Desktop/primeira/Gabarito Prateleira - SIAD - 202409.xlsx')


#da seplag:
#df_apuracao = pd.read_excel('C:/Users/X15848318638/Desktop/primeira/Apuração do faturamento 202308.xlsx', sheet_name = 'Apuração do Faturamento')
#df_gabarito = pd.read_excel('C:/Users/X15848318638/Desktop/primeira/Gabarito Órgão-Sigla 202406.xlsx')
#df_cod_orgao = pd.read_excel('C:/Users/X15848318638/Desktop/primeira/Gabarito Prateleira - SIAD - 202409.xlsx')

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
df_apuracao['Observação'] = 'Sem observação'


# Etapa Preencheendo  a coluna “Exercício” com o ano referente ao da coluna Data da Aprovação
df_apuracao['Exercício'] = df_apuracao['Data da Aprovação'].dt.year.fillna(0).astype(int)


# Etapa de exclusão das linhas do ano anterior
## supondo que o usuário colocou o ano de 2023
#df_apuracao = df_apuracao['Data do Fato gerador']==

# Etapa de exclusão de valores "Não se aplica"
df_apuracao['Data limite de entrega - pedido original'] = df_apuracao['Data limite de entrega - pedido original'].replace('Não se aplica', "")
df_apuracao['Dias de atraso - pedido original'] = df_apuracao['Dias de atraso - pedido original'].replace('Não se aplica', "")
df_apuracao['Data limite de entrega - entrega corretiva'] = df_apuracao['Data limite de entrega - entrega corretiva'].replace('Não se aplica', "")
df_apuracao['Dias de atraso - entrega corretiva'] = df_apuracao['Dias de atraso - entrega corretiva'].replace('Não se aplica', "")

import pandas as pd

# Supondo que você já tenha um DataFrame criado
# df = pd.DataFrame(...)

# Definindo o caminho completo onde o arquivo será salvo
caminho_do_arquivo = '/caminho/para/pasta/nome_do_arquivo.xlsx'

# Salvando o DataFrame em um arquivo Excel com um nome específico para a planilha
df_apuracao.to_excel('C:/Users/cante/OneDrive/Desktop/primeira/Valor Faturado.xlsx',index=False,sheet_name='Valor Faturado')

df_apuracao.info()
#print(df_apuracao['Data limite de entrega - pedido original'])




#df_apuracao.info()
#print(df_apuracao)

