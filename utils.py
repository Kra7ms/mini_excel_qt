import re

def cell_to_index(cell: str):
    match = re.match(r"([A-Z]+)([0-9]+)", cell.upper())
    if not match:
        return None
    
    col_letters, row_number = match.groups()

    col = 0
    for ch in col_letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)

    row = int(row_number) -1
    return row, col -1
    
def index_to_cell(row: int, col: int):
    name = ""
    col += 1
    while col:
        col, rem = divmod(col - 1, 26)
        name = chr(rem + ord("A")) + name
    return f"{name}{row + 1}"
