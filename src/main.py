import datetime
import json
import requests
from itertools import zip_longest

import PySimpleGUI as sg
import xlwings as xw

# .env
env_json = json.load(open("../.env", 'r', encoding="utf-8_sig"))

# excelOpen
book = xw.Book("./Record.xlsx")


def read_row_xl(sheet, first_cell_address):
    time_list = []
    row = first_cell_address[0]
    col = first_cell_address[1]
    while sheet.cells(row, col).value is not None:
        time_list.append(sheet.cells(row, col).value)
        row += 1
    return time_list


def refresh_2d_list_from_xl():
    sheet = book.sheets['Time']
    start_time_list = read_row_xl(sheet, [2, 1])
    end_time_list = read_row_xl(sheet, [2, 2])
    time_list = list(zip_longest(start_time_list, end_time_list))

    if len(time_list) != 0:
        window['table'].update(time_list)
    book.save()


def append_last_row_cell(sheet, col_address, data):
    row_address = 1
    while sheet.cells(row_address, col_address).value is not None:
        row_address += 1
    sheet.cells(row_address, col_address).value = data


def send_to_discord(message):
    webhook_url = env_json["web_hook_url"]
    main_content = {'content': message}
    headers = {'Content-Type': 'application/json'}
    requests.post(webhook_url, json.dumps(main_content), headers=headers)


# pySimpleGui setting
sg.change_look_and_feel('GreenMono')

layout = [[
    sg.Table(
        [[], []],
        headings=['start', 'end'],
        auto_size_columns=False,
        col_widths=[15, 15],
        justification='left',
        text_color='#000000',
        background_color='#cccccc',
        alternating_row_color='#ffffff',
        header_text_color='#0000ff',
        header_background_color='#cccccc',
        key='table'
    )],
    [sg.Button('start', key='startBtn', size=(15, 2)), sg.Button('end', key='endBtn', size=(15, 2))]]

window = sg.Window('Time Stamp', layout, resizable=True, finalize=True)

refresh_2d_list_from_xl()

# Event loop
while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED:
        break

    elif event == 'startBtn':
        time_sheet = book.sheets['Time']

        now = datetime.datetime.now()

        append_last_row_cell(time_sheet, 1, now)

        refresh_2d_list_from_xl()

        send_to_discord("clock-in:" + now.strftime('%m-%d %H:%M'))

    elif event == 'endBtn':
        time_sheet = book.sheets['Time']

        now = datetime.datetime.now()

        append_last_row_cell(time_sheet, 2, now)

        refresh_2d_list_from_xl()

        send_to_discord("clock-out:" + now.strftime('%m-%d %H:%M'))

book.close()
window.close()
