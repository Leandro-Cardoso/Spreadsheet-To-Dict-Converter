# Conversor de planilha para dicionario

Recebe um arquivo do tipo **xlsx**,  **xls**, Excel ou arquivos **CSV** e converte para um dicionário. Identifica automaticamente os dados de acordo com a posição dos cabeçalhos. Por sua vez, os dados podem ser salvos em **JSON** ou em um banco de dados.

## Funções:

* spreadsheet_to_dict(**parametros**):

    Recebe um arquivo do tipo xlsx, xls, Excel ou arquivos CSV e converte para um dicionário. Identifica automaticamente os dados de acordo com a posição dos cabeçalhos.

    ### Parametros:

    * file_path:

        Caminho completo da planilha, incluindo a extensão.

    * ws_name:

        Nome da aba da planilha.

    * horizontal_headers:

        Lista com os nomes dos cabeçalhos dos quais os dados estão organizados na direção horizontal.

    * vertical_headers:
    
        Lista com os nomes dos cabeçalhos dos quais os dados estão organizados na direção vertical.

    ### Retornos:

    * dict:

        Retorna um dicionário com a estrutura de chave com os cabeçalhos da planilha e o valor de cada chave é uma lista com os valores relacionados ao cabeçalho.
