import time
import json
import os
from pprint import pprint
from datetime import datetime, timedelta
from dateutil import parser
from multiprocessing import Process
import pandas as pd
import numpy as np


def save_excel_multiple_sheets(data: dict[pd.DataFrame], name):
    writer = pd.ExcelWriter(name)
    for key in data:
        data[key].to_excel(writer, key, index=False)
    writer.save()


def process_report(report: pd.DataFrame, refs: dict, client_set: set, entities: dict):
    report = report.applymap(lambda x: x.rstrip() if type(x) == str else x)
    look_for_curr_tl = 'Mevcut Bakiye:'
    look_for_curr_en = 'Current Balance:'
    offset_curr = 2
    curr_find = np.where(report == look_for_curr_tl)
    if len(curr_find[0]):
        trans = {
            'transaction_time': 'Tarih/Saat',
            'value_date': 'Valör',
            'sum': 'İşlem Tutarı*',
            'balance_after': 'Bakiye',
            'comment': 'Açıklama',
            'reference': 'Referans'
        }
        look_for_start = 'Tarih/Saat'
    else:
        curr_find = np.where(report == look_for_curr_en)
        trans = {
            'transaction_time': 'Date/Time',
            'value_date': 'Value Date',
            'sum': 'Transaction\nAmount*',
            'balance_after': 'Balance',
            'comment': 'Description',
            'reference': 'Reference'
        }
        look_for_start = 'Date/Time'
    ent = None
    for e in entities:
        e_find = np.where(report == f'Dear {entities[e]}')
        if len(e_find[0]):
            ent = e
            break
    if not ent:
        find_dear = report.applymap(lambda x: x[:4] if type(x) == str else '')
        dear_find = np.where(find_dear == 'Dear')
        if len(dear_find[0]):
            return None, 'new'
        return None, 'none'
    curr_y, curr_x = curr_find[0][0], curr_find[1][0] + offset_curr
    curr = report.iloc[curr_y, curr_x].split()[1]
    start = np.where(report == look_for_start)[0][0]
    start_line = report.iloc[start]
    cols = {}
    for c in trans:
        cols[c] = np.where(start_line == trans[c])[0][0]

    res = []
    while True:
        start += 1
        line = {c: report.iloc[start, cols[c]] for c in trans}
        if not line['transaction_time']:
            break
        if line['reference'] in refs[ent]:
            continue
        refs[ent].add(line['reference'])
        line['currency'] = curr
        line['value_date'] = str(datetime.strptime(line['value_date'], '%d/%m/%Y'))[:10]
        line['client'] = 0
        for cl in client_set:
            if cl in line['comment']:
                line['client'] = cl
                break
        res.append(line)

    return res, ent
