import os
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# definindo enderecos dos arquivos utilizados
fonte = './Data.xlsx'
destino = './consolidated_data.xlsx'

# funções de uso geral
def abrir_ou_criar_planilha(caminho_arquivo: str, nome_aba: str):
    # verifica se o arquivo existe
    if os.path.exists(caminho_arquivo):
        wb = load_workbook(caminho_arquivo)
    else:
        wb = Workbook()
        # remove a aba padrão se for criar um novo arquivo
        if 'Sheet' in wb.sheetnames:
            default_sheet = wb['Sheet']
            wb.remove(default_sheet)

    # verifica se a aba desejada já existe
    if nome_aba in wb.sheetnames:
        ws = wb[nome_aba]
    else:
        ws = wb.create_sheet(title=nome_aba)

    return wb, ws

def inserir_df_na_aba(ws: any, df: pd.DataFrame, incluir_header: bool = True, incluir_index: bool = True):
    """
    ws: worksheet do openpyxl
    df: pandas DataFrame
    incluir_header (bool): inclui nomes das colunas
    incluir_index (bool): inclui índice do DataFrame
    """
    # limpando possiveis dados existentes na aba
    ws.delete_rows(1, ws.max_row)

    # iterando sobre as linhas e colunas do df e inserindo os valores célula a célula na planilha
    for r_idx, row in enumerate( # linhas
        dataframe_to_rows(df, index=incluir_index, header=incluir_header),
        start=1 # index da linha onde os dados começarão a ser inseridos, o padrão do enumerate é 0, diferindo da indexação do Excel que começa em 1
    ):
        for c_idx, value in enumerate(row, start=1): # colunas
            ws.cell(row=r_idx, column=c_idx, value=value)

# teste 1: consolidar os dados de cada aba em um único arquivo
def teste_1():
    todas_abas = pd.read_excel(fonte, sheet_name=None)
    aux = []

    # gerando um df com os dados de cada aba, alterando a coluna "Preço" para o nome da aba/papel correspondente
    for nome_aba, df in todas_abas.items():
        df = df.rename(columns={'Preço': f'{nome_aba}'})
        aux.append(df)

    # pegando os dados da primeira aba para iniciar a consolidação
    infos_consolidadas = aux[0] 
    for df in aux[1:]:
        # unindo os dados dados das abas seguintes, utilizando a coluna "Data" como chave de junção
        infos_consolidadas = pd.merge(infos_consolidadas, df, on='Data', how='inner')

    # definindo a coluna "Data" como índice para facilitar o acesso aos 
    # infos_consolidadas = infos_consolidadas.set_index('Data')

    # guardando a consolidação no excel de destino
    wb, ws = abrir_ou_criar_planilha('./consolidated_data1.xlsx', 'consolidated_data')
    inserir_df_na_aba(ws, infos_consolidadas,incluir_index=False)
    wb.save('./consolidated_data1.xlsx')

    # teste 1.1: calcular o retorno diário de cada papel
    def teste_1_1():
        # lendo o arquivo consolidado
        df_precos = pd.read_excel(destino) 

        # definindo a coluna "Data" como índice para facilitar o cálculo do retorno diário
        df_precos = df_precos.set_index('Data') 
        
        # calculando os retornos e adicionando prefixos às colunas facilitar na concatenção dos df
        df_retornos = df_precos.pct_change().add_prefix('Retorno_')

        # unindo retornos e precos em um único df
        df_final = pd.concat([df_precos, df_retornos], axis=1)

        # organizando as colunas para que os retornos fiquem ao lado dos preços correspondentes
        colunas_organizadas = []
        for coluna in df_precos.columns:
            colunas_organizadas.append(coluna)
            colunas_organizadas.append(f'Retorno_{coluna}')

        df_final = df_final[colunas_organizadas]

        return df_final
    # infos_consolidadas = teste_1_1()

    # infos_consolidadas.to_excel(destino, index=True) # guardando o df final com preços e retornos em um excel

# def teste_2():


teste_1()