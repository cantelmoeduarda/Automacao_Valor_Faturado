#alteração das planilhas sem o botão
import pandas as pd
import openpyxl
#do pessoal:
df_apuracao = pd.read_excel('C:/Users/cante/OneDrive/Desktop/primeira/Apuração do faturamento 202408.xlsx', sheet_name = 'Apuração do Faturamento')
df_orgao_sigla = pd.read_excel('C:/Users/cante/OneDrive/Desktop/primeira/Gabarito Órgão-Sigla 202409.xlsx')
df_prateleira = pd.read_excel('C:/Users/cante/OneDrive/Desktop/primeira/Gabarito Prateleira - SIAD - 202409.xlsx')

def alteracao_planilha():
    # Etapa de criação da coluna Ajustes e 0 (zero) em toda coluna
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


    #Etapa de tirar o time da data
    obj_to_data = ['Data limite de entrega - entrega corretiva', 'Data limite de entrega - pedido original']

    for coluna in obj_to_data:
        df_apuracao[coluna] = pd.to_datetime(df_apuracao[coluna])
        
    colunas_de_data = ['Data do Fato Gerador', 'Data da Aprovação', 'Data do Ateste','Data limite de entrega - entrega corretiva', 'Data limite de entrega - pedido original']
    for data in colunas_de_data:
        df_apuracao[data] = pd.to_datetime(df_apuracao[data]).dt.strftime('%d/%m/%Y')
        
    df_apuracao.to_excel('C:\\Users\\cante\\OneDrive\\Desktop\\primeira\\Quase_la.xlsx',sheet_name='Valor Faturado',index=False)
        
    #da seplag:
    #df_apuracao = pd.read_excel('C:/Users/X15848318638/Desktop/primeira/Apuração do faturamento 202308.xlsx', sheet_name = 'Apuração do Faturamento')
    #df_orgao_sigla = pd.read_excel('C:/Users/X15848318638/Desktop/primeira/Gabarito Órgão-Sigla 202406.xlsx')
    #df_prateleira = pd.read_excel('C:/Users/X15848318638/Desktop/primeira/Gabarito Prateleira - SIAD - 202409.xlsx')


    # Etapa de criação da coluna Ajustes e 0 (zero) em toda coluna
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


    #Etapa de tirar o time da data
    obj_to_data = ['Data limite de entrega - entrega corretiva', 'Data limite de entrega - pedido original']

    for coluna in obj_to_data:
        df_apuracao[coluna] = pd.to_datetime(df_apuracao[coluna])
        
    colunas_de_data = ['Data do Fato Gerador', 'Data da Aprovação', 'Data do Ateste','Data limite de entrega - entrega corretiva', 'Data limite de entrega - pedido original']
    for data in colunas_de_data:
        df_apuracao[data] = pd.to_datetime(df_apuracao[data]).dt.strftime('%d/%m/%Y')
        
    df_apuracao.to_excel('C:\\Users\\cante\\OneDrive\\Desktop\\primeira\\Quase_la.xlsx',sheet_name='Valor Faturado',index=False)
