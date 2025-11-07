from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

import re

from utils import is_valid_file_path, has_duplicates, get_uppercase_list

#|--------------------------------------------------------------|
#| INTERNAL FUNCTIONS:
#|--------------------------------------------------------------|
def _get_cel_coordinate(ws, value) -> str:
    '''
    Pegar a posição de uma determinada celula na tabela.
    '''
    for row in ws.iter_rows():
        for cel in row:
            if cel.value == value:
                return cel.coordinate

    raise FileNotFoundError('Valor da celula não encontrado na tabela.')

def _is_col_none(ws, col_letter: str) -> bool:
    '''
    Verifica se a coluna está vazia.
    '''
    for row in range(1, ws.max_row + 1):
        if ws[f'{col_letter}{row}'].value:
            return False
    
    return True

def _is_row_none(ws, row: int) -> bool:
    '''
    Verifica se a linha está vazia.
    '''
    for col_index in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_index)

        if ws[f'{col_letter}{row}'].value:
            return False
    
    return True

def _get_row_values(ws, header_value, stop_values: list) -> list:
    '''
    Pega todos os valores de uma linha a partir de um cabeçalho.
    '''
    cel_header = ws[_get_cel_coordinate(ws, header_value)]
    col_values = []

    for col_index in range(cel_header.column + 1, ws.max_column + 1): 
        col_letter = get_column_letter(col_index)
        col_value = ws[f'{col_letter}{cel_header.row}'].value

        if _is_col_none(ws, col_letter):
            continue
        
        if col_value in stop_values:
            break
        
        col_values.append(col_value)

    return col_values

def _get_col_values(ws, header_value, stop_values: list) -> list:
    '''
    Pega todos os valores de uma coluna a partir de um cabeçalho.
    '''
    cel_header = _get_cel_coordinate(ws, header_value)
    col_letter = re.sub(r'\d', '', cel_header)
    col_number = int(re.sub(r'[^0-9]', '', cel_header))
    col_values = []

    for row in range(col_number + 1, ws.max_row + 1):
        cel_value = ws[f'{col_letter}{row}'].value

        if _is_row_none(ws, row):
            continue

        if cel_value in stop_values:
            break

        col_values.append(cel_value)

    return col_values

#|--------------------------------------------------------------|
#| MAIN FUNCTIONS:
#|--------------------------------------------------------------|
def spreadsheet_to_dict(
        file_path: str,
        ws_name: str,
        horizontal_headers: list = [],
        vertical_headers: list = []
) -> dict:
    '''
    Recebe um arquivo do tipo xlsx, xls, Excel ou arquivos CSV e converte para um dicionário. Identifica automaticamente os dados de acordo com a posição dos cabeçalhos.
    '''
    allowed_extensions = ['xlsx', 'xls', 'csv']
    new_dict = {}

    # Validações:
    if not is_valid_file_path(file_path, allowed_extensions):
        raise FileNotFoundError('Caminho do arquivo ou extensão invalida.')
    
    if not horizontal_headers and not vertical_headers:
        raise ValueError('É necessario informar ao menos uma das listas de cabeçalhos (headers).')
    
    if has_duplicates(horizontal_headers + vertical_headers):
        raise ValueError('Não pode haver mais de um cabeçalho (header) com o mesmo nome.')
    
    # Buscar dados:
    wb = load_workbook(file_path, data_only=True)
    ws = wb[ws_name]

    if vertical_headers:
        # Dados em linhas:
        for header in vertical_headers:
            new_dict[str(header).upper()] = _get_row_values(ws, header, vertical_headers)

    if horizontal_headers:
        # Dados em colunas:
        stop_values = horizontal_headers + vertical_headers

        for header in vertical_headers:
            key = str(header).upper()
            
            if key in new_dict and isinstance(new_dict[key], list):
                if not None in new_dict[key]:
                    stop_values.extend(new_dict[key])
        
        for header in horizontal_headers:
            new_dict[str(header).upper()] = _get_col_values(ws, header, stop_values)

    # Prerações:
    for header in new_dict:
        new_dict[header] = get_uppercase_list(new_dict[header])

    return new_dict

#|--------------------------------------------------------------|
#| TESTS:
#|--------------------------------------------------------------|
if __name__ == '__main__':
    # Configurações:
    file_path = 'C:\\Users\\leand\\Documents\\PROJETOS\\spreadsheet-to-dict-converter\\TESTE.xlsx'
    sheet_name = 'teste'
    horizontal_headers = ['Nome', 'Valor']
    vertical_headers = ['Total']

    # Execução:
    dict = spreadsheet_to_dict(file_path, sheet_name, horizontal_headers, vertical_headers)

    print(dict)
