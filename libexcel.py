import pickle
import requests
import os
import json


def calculate_percentage(buy, sell):
    return (sell - buy) / buy


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


def load_value(s, r, c):
    return load_cell(s, r, c).value


def set_cell(s, r, c, value):
    cell = s.cell(row=r, column=c)
    cell.value = value


def column_number(s):
    ret = 0
    s = s.upper()
    for i in range(0, len(s)):
        ret = ret * 26 + ord(s[i]) - ord('A') + 1
    return ret


def column_idx(char):
    return ord(char.upper()) - ord('A') + 1


def get_history(symbol, unix_sod, unix_today):
    filepath = 'history/' + symbol + '.pickle'
    filepath2 = 'history/' + symbol + '.txt'

    if os.path.isfile(filepath):
        with open(filepath, 'rb') as f:
            history = pickle.load(f)
        return history
    else:
        url = f'https://ws.api.cnyes.com/charting/api/v1/history?resolution=D&symbol=TWS:{symbol}:STOCK&from={unix_today}&to={unix_sod}&quote=1'

        print(url)
        r = requests.get(url)
        # print(r.content)
        data = json.loads(r.content)
        if data['statusCode'] != 200:
            print("HTTP request error!")
            return None

        history = []

        for t in data['data']['t']:
            # print(datetime.datetime.utcfromtimestamp(t).strftime('%Y-%m-%d'))
            pass

        T = data['data']['t']
        O = data['data']['o']
        H = data['data']['h']
        L = data['data']['l']
        C = data['data']['c']
        V = data['data']['v']

        with open(filepath2, 'w') as f:
            for t, o, h, l, c, v in zip(T, O, H, L, C, V):
                history.insert(0, (t, o, h, l, c, v))
                f.write(f'{T} {O} {H} {L} {C} {V}\n')

        with open(filepath, 'wb') as f:
            pickle.dump(history, f)

        return get_history(symbol, unix_sod, unix_today)
