#|--------------------------------------------------------------|
#| DEPENDENCIAS:
#|--------------------------------------------------------------|
# is_valid_file_path():
from pathlib import Path

#|--------------------------------------------------------------|
#| UTILS EM ORDEM ALFABETICA:
#|--------------------------------------------------------------|
def get_uppercase_list(data_list: list) -> list:
    '''
    Passa os elementos de uma lista para maiusculas.
    '''
    uppercase_list = [
        str(element).upper() if isinstance(element, str) else element 
        for element in data_list
    ]

    return uppercase_list

def has_duplicates(list: list) -> bool:
    '''
    Verifica se existe pelo menos um elemento duplicado em uma lista.
    '''
    if not list:
        return False
    
    list = [str(element) for element in list]
    
    return len(list) != len(set(list))

def is_valid_file_path(
        file_path: str,
        exts: list[str] = None
) -> bool:
    '''
    Verifica se o caminho do arquivo é valido e se corresponde a extensão nescessária.
    '''
    p = Path(file_path)

    if not p.is_file():
        return False

    if exts:
        file_extension = p.suffix.lower().replace('.', '')
        allowed_extensions = {ext.lower() for ext in exts}

        if file_extension not in allowed_extensions:
            return False

    return True
