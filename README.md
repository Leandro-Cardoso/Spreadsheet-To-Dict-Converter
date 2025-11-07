# Conversor de planilha para dicionario

Recebe um arquivo do tipo **xlsx**,  **xls**, Excel ou arquivos **CSV** e converte para um dicionário. Identifica automaticamente os dados de acordo com a posição dos cabeçalhos. Por sua vez, os dados podem ser salvos em **JSON** ou em um banco de dados.

## Funções:

### spreadsheet_cols_to_dict():

    Recebe um arquivo do tipo xlsx, xls, Excel ou arquivos CSV e converte para um dicionário. Identifica automaticamente os dados de acordo com a posição dos cabeçalhos.

    ### Parametros:

    * file_path:

        Caminho completo da planilha, incluindo a extensão.

    * ws_name:

        Nome da aba da planilha.

    * headers:

        Lista com os nomes dos cabeçalhos das colunas.

    * end_row:
    
        Ultima linha a ser considerada das colunas.

    ### Retornos:

    * dict:

        Retorna um dicionário com a estrutura de chave com os cabeçalhos da planilha e o valor de cada chave é uma lista com os valores relacionados ao cabeçalho.

### spreadsheet_rows_to_dict():

    Recebe um arquivo do tipo xlsx, xls, Excel ou arquivos CSV e converte para um dicionário. Identifica automaticamente os dados de acordo com a posição dos cabeçalhos.

    ### Parametros:

    * file_path:

        Caminho completo da planilha, incluindo a extensão.

    * ws_name:

        Nome da aba da planilha.

    * headers:

        Lista com os nomes dos cabeçalhos das linhas.

    * end_col:
    
        Ultima coluna a ser considerada das linhas.

    ### Retornos:

    * dict:

        Retorna um dicionário com a estrutura de chave com os cabeçalhos da planilha e o valor de cada chave é uma lista com os valores relacionados ao cabeçalho.
