def convert_to_str(input):
    if type(input) is str:
        return input.rstrip(' ').lstrip(' ').rstrip(' ')

    if type(input) is int:
        return str(input)

    return input


def load_cell(s, r, c):
    x = s.cell(row=r, column=c)
    x.value = convert_to_str(x.value)
    return x


def set_cell(s, r, c, value):
    cell = s.cell(row=r, column=c)
    cell.value = value


def column_number(s):
    ret = 0
    s = s.upper()
    for i in range(0, len(s)):
        ret = ret * 26 + ord(s[i]) - ord('A') + 1
    return ret
