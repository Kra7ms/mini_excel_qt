def cell_to_index(ref: str):
    col = 0
    i = 0
    while i < len(ref) and ref[i].isalpha():
        col = col * 26 + (ord(ref[i]) - ord("A") + 1)
        i += 1
    row = int(ref[i:]) - 1
    return row, col - 1


def index_to_cell(row: int, col: int):
    col += 1
    name = ""
    while col:
        col, rem = divmod(col - 1, 26)
        name = chr(ord("A") + rem) + name
    return f"{name}{row + 1}"

def expand_range(start_ref: str, end_ref: str):
    """
    "A1", "C3"  ->  ["A1","A2","A3","B1","B2","B3","C1","C2","C3"]
    """
    start = cell_to_index(start_ref)
    end = cell_to_index(end_ref)

    if not start or not end:
        return []

    r1, c1 = start
    r2, c2 = end

    cells = []

    for r in range(min(r1, r2), max(r1, r2) + 1):
        for c in range(min(c1, c2), max(c1, c2) + 1):
            cells.append(index_to_cell(r, c))

    return cells