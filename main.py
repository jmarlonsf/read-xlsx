import argparse
import pandas as pd
import re


def count_line_cell_notnull(line):
    """
    Conta o número de células não nulas em uma linha.

    Parameters:
        line (pandas.Series): Uma linha do DataFrame.

    Returns:
        int: Número de células não nulas na linha.
    """
    return line.count()


def clean_title(title):
    """
    Limpa e formata um título para ser usado como nome de arquivo.

    Parameters:
        title (str): Título a ser limpo.

    Returns:
        str: Título limpo e formatado.
    """
    title = title.lower()
    title = re.sub(r'\s+', ' ', title)
    title = title.replace(' ', '_')
    title = re.sub(r'[^a-zA-Z0-9áéíóúÁÉÍÓÚâêîôÂÊÎÔãõÃÕàèìòÀÈÌÒçÇ _]', '', title)
    return title


def save_dataframe(dataframe, title):
    """
    Salva um DataFrame como um arquivo CSV, se não estiver vazio.

    Parameters:
        dataframe (pandas.DataFrame): DataFrame a ser salvo.
        title (str): Título para nomear o arquivo.

    Returns:
        pandas.DataFrame, list: DataFrame vazio e lista vazia.
    """
    if not dataframe.empty and not dataframe.columns.empty:
        title = clean_title(title)
        path = f"./data/{title}.csv"
        dataframe = dataframe.dropna(axis=1, how='all')
        dataframe.to_csv(path, index=False)
        print(f"Arquivo salvo em {path}")
        return pd.DataFrame(), []

    return dataframe, []


def extract(arquivo_excel):
    """
    Extrai dados de tabelas de um arquivo Excel e salva como arquivos CSV.

    Parameters:
        arquivo_excel (str): Caminho para o arquivo Excel a ser processado.
    """
    try:
        # Usando a função ExcelFile para abrir o arquivo Excel
        xls = pd.ExcelFile(arquivo_excel)

        # Obtendo uma lista de nomes de todas as abas no arquivo
        nomes_das_abas = xls.sheet_names

        # Criando um dicionário para armazenar todos os DataFrames das abas
        planilha = {}

        # Iterando sobre os nomes das abas e lendo os dados em DataFrames
        for aba in nomes_das_abas:

            planilha[aba] = pd.read_excel(arquivo_excel, sheet_name=aba, header=None)

            dataframe = pd.DataFrame()
            title = ""
            columns = []

            for idx, line in planilha[aba].iterrows():
                line_size = count_line_cell_notnull(line)

                if line_size == 0:
                    # chamando o método abaixo para caso pule de uma tabela para outra
                    dataframe, columns = save_dataframe(dataframe=dataframe, title=title)
                    continue
                elif line_size == 1:
                    dataframe, columns = save_dataframe(dataframe=dataframe, title=title)
                    # Esse fluxo considera que o texto anterior a tabela dará nome ao arquivo
                    title = f"{aba}_{line.dropna().values[0]}"
                else:
                    if dataframe.empty and dataframe.columns.empty:
                        columns = line.values
                        dataframe = pd.DataFrame(columns=columns)
                    else:
                        new_line = pd.Series(line.values, index=columns)
                        dataframe = pd.concat([dataframe, new_line.to_frame().T], ignore_index=True)

            save_dataframe(dataframe=dataframe, title=title)

    except Exception as e:
        print(f"Erro durante a execução: {str(e)}")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Process xlsx files')
    parser.add_argument('--path', type=str, help='file that will be analyzed', required=True)

    args = parser.parse_args()

    extract(args.path)
