# read-xlsx

## Descrição

O código é projetado para analisar arquivos XLSX que contêm tabelas em diferentes abas. Essas tabelas devem ser separadas por pelo menos uma linha vazia e devem conter uma linha com uma célula anterior à tabela que será usada como nome dela.

## Instalação das Bibliotecas

Antes de executar o código, certifique-se de ter as bibliotecas necessárias instaladas. Você pode instalá-las executando os seguintes comandos:

```bash
pip install pandas==2.1.3
pip install openpyxl==3.1.2
pip install argparse==1.4.0
```

Estes comandos instalam as versões específicas das bibliotecas Pandas, OpenPyXL e argparse necessárias para a execução do código.

## Arquivo de Execução

Para executar o código, utilize o seguinte comando no terminal, substituindo `arquivo_de_origem.xlsx` pelo caminho do seu arquivo XLSX:

```bash
python main.py --path ./arquivo_de_origem.xlsx
```

Isso iniciará a execução do programa principal.

## Diretório de Destino

Os resultados gerados pelo programa serão salvos no diretório `./data`. Certifique-se de que este diretório existe antes de executar o programa. Caso contrário, você pode criar manualmente usando o seguinte comando:

```bash
mkdir data
```

Este comando cria o diretório `data` na mesma pasta onde o arquivo `main.py` está localizado. Os resultados do programa serão armazenados neste diretório.
