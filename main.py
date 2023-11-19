import argparse
import pandas as pd
import re


def count_line_cell_notnull(line):
    return line.count()


def group_and_clean_dataframe(title="", dataframe=pd.DataFrame(), columns=list):
    if columns is None or len(columns) != 0:
        columns = []

    if not dataframe.empty:
        # apaga todas as colunas vazias
        dataframe = dataframe.dropna(axis=1, how='all')
        title = title.lower()
        title = re.sub(r'\s+', ' ', title)
        title = title.replace(' ', '_')
        title = re.sub(r'[^a-zA-Z0-9áéíóúÁÉÍÓÚâêîôÂÊÎÔãõÃÕàèìòÀÈÌÒçÇ _]', '', title)
        path = f"./data/{title}.csv"
        dataframe.to_csv(path, index=False)
        print(f"Arquivo salvo em {f"./data/{title}.csv"}")

    dataframe = pd.DataFrame()

    return dataframe, columns


def extract(arquivo_excel):
    # Use a função ExcelFile para abrir o arquivo Excel
    xls = pd.ExcelFile(arquivo_excel)

    # Obtenha uma lista de nomes de todas as abas no arquivo
    nomes_das_abas = xls.sheet_names

    # Criando um dicionário para armazenar todos os DataFrames das abas
    planilha = {}

    # Iterando sobre os nomes das abas e leendo os dados em DataFrames
    for aba in nomes_das_abas:

        planilha[aba] = pd.read_excel(arquivo_excel, sheet_name=aba, header=None)

        dataframe = pd.DataFrame()
        title = ""
        columns = []
        for idx, line in planilha[aba].iterrows():
            line_size = count_line_cell_notnull(line)
            if line_size == 0:
                #chamando o método abaixo para caso pule de uma tabela para outra
                dataframe, columns = group_and_clean_dataframe(title=title, dataframe=dataframe,
                                                               columns=columns)
                continue
            elif line_size == 1:
                # Esse fluxo considera que o texto anterior a tabela dará nome ao arquivo
                title = f"{aba}_{line.dropna().values[0]}"
            else:
                if dataframe.empty and dataframe.columns.empty:
                    columns = line.values
                    dataframe = pd.DataFrame(columns=columns)
                else:
                    new_line = pd.Series(line.values, index=columns)
                    dataframe = pd.concat([dataframe, new_line.to_frame().T], ignore_index=True)

        group_and_clean_dataframe(title=title, dataframe=dataframe, columns=columns)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Process xlsx files')
    parser.add_argument('--path', type=str,
                        help='file that will be analized', required=True)

    args = parser.parse_args()

    extract(args.path)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
