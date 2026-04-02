import pandas as pd

fonte = './Data.xlsx'

def teste_1():
    todas_abas = pd.read_excel(fonte, sheet_name=None)
    aux = []

    for nome_aba, df in todas_abas.items():
        df = df.rename(columns={'Preço': f'{nome_aba}'})
        aux.append(df)

    infos_consolidadas = aux[0]
    for df in aux[1:]:
        infos_consolidadas = pd.merge(infos_consolidadas, df, on='Data', how='inner')

    print(infos_consolidadas)

teste_1()