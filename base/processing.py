import time
import json
import os
from pprint import pprint
from datetime import datetime, timedelta
from dateutil import parser
from multiprocessing import Process
import pandas as pd
import numpy as np
import shutil
from pathlib import Path
try:
    import win32com.client as win32
    xl_app = win32.DispatchEx("Excel.Application")
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
except Exception as ex:
    print("No Windows module")
    xl_app = None
    outlook = None
working_directory = os.path.dirname(os.path.realpath(__file__))


def save_excel_multiple_sheets(data: dict, name):
    writer = pd.ExcelWriter(name)
    for key in data:
        data[key].to_excel(writer, key, index=False)
    writer.save()


def process_report(report: pd.DataFrame, refs: dict, client_set: set, entities: dict, args: dict):
    report = report.applymap(lambda x: x.rstrip() if type(x) == str else x)
    curr_find = np.where(report == args['look_for_curr']['tur'])
    if len(curr_find[0]):
        trans = args['trans_tur']
        dear = args['look_for_client']['tur']
    else:
        curr_find = np.where(report == args['look_for_curr']['eng'])
        trans = args['trans_eng']
        dear = args['look_for_client']['eng']
    look_for_start = trans['transaction_time']
    ent = None
    for e in entities:
        e_find = np.where(report == f'{dear} {entities[e]}')
        if len(e_find[0]):
            ent = e
            break
    if not ent:
        find_dear = report.applymap(lambda x: x[:4] if type(x) == str else '')
        dear_find = np.where(find_dear == dear)
        if len(dear_find[0]):
            return None, 'new'
        return None, 'none'
    curr_y, curr_x = curr_find[0][0], curr_find[1][0] + args['look_for_curr']['horizontal_offset']
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
