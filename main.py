import pandas as pd

# definindo arquivo com dados brutos
fonte = './Data.xlsx'
destino = './consolidated_data.xlsx'


# teste 1: consolidar os dados de cada aba em um único arquivo
def teste_1():
    todas_abas = pd.read_excel(fonte, sheet_name=None)
    aux = []

    # gerando um df com os dados de cada aba, alterando a coluna "Preço" para o nome da aba/papel correspondente
    for nome_aba, df in todas_abas.items():
        df = df.rename(columns={'Preço': f'{nome_aba}'})
        aux.append(df)

    infos_consolidadas = aux[0] # pegando os dados da primeira aba para iniciar a consolidação
    for df in aux[1:]:
        # unindo os dados dados das abas seguintes, utilizando a coluna "Data" como chave de junção
        infos_consolidadas = pd.merge(infos_consolidadas, df, on='Data', how='inner')

    # guardando a consolidação em um excel
    infos_consolidadas.to_excel(destino, index=False)

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
    infos_consolidadas = teste_1_1()

    infos_consolidadas.to_excel(destino, index=True) # guardando o df final com preços e retornos em um excel

# def teste_2():


teste_1()