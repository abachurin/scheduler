import time
import json
import os
from pprint import pprint
from datetime import datetime, timedelta
from dateutil import parser
from multiprocessing import Process
import pandas as pd
import numpy as np


def process_report(report: pd.DataFrame, refs: set, client_set: set):
    trans = {
        'transaction_time': 'Tarih/Saat',
        'value_date': 'Valör',
        'sum': 'İşlem Tutarı*',
        'balance_after': 'Bakiye',
        'comment': 'Açıklama',
        'reference': 'Referans'
    }
    look_for_curr = 'Mevcut Bakiye:'
    offset_curr = 2
    look_for_start = 'Tarih/Saat'

    curr_find = np.where(report == look_for_curr)
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
        if line['reference'] in refs:
            continue
        line['currency'] = curr
        line['client'] = None
        for cl in client_set:
            if cl in line['comment']:
                line['client'] = cl
                break
        res.append(line)

    return res
