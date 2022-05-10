import os

from base.processing import *


def func_1(args):
    path = args.get('path', 'no path in config file')
    try:
        wb = xl_app.Workbooks.Open(path)
        wb.RefreshAll()
        xl_app.CalculateUntilAsyncQueriesDone()
        xl_app.DisplayAlerts = False
        wb.Save()
        wb.Close()
        print(f'{datetime.now()}: file {path} was refreshed')
    except Exception as ex:
        print(f'failed to refresh {path}')
        print(ex)


def func_2(args):
    path = args.get('path', 'no path in config file')
    try:
        wb = xl_app.Workbooks.Open(path)
        wb.RefreshAll()
        xl_app.CalculateUntilAsyncQueriesDone()
        xl_app.DisplayAlerts = False
        wb.Save()
        wb.Close()
        print(f'{datetime.now()}: file {path} was refreshed')
    except Exception as ex:
        print(f'failed to refresh {path}')
        print(ex)


def func_3(args):
    print(args['to_print'])


def iz_consolidated(args):
    target_directory = args["target_directory"]
    Path(target_directory).mkdir(parents=True, exist_ok=True)
    consolidated_file = args['consolidated_file']
    memory_file = args["memory_file"]
    try:
        with open(memory_file, "r", encoding='utf-8') as f:
            memory = json.load(f)
    except Exception:
        memory = []
    new_files = [v for v in os.listdir(target_directory) if (not v.startswith('.') and v not in memory)]
    if not new_files:
        print(f'No files to process in {target_directory}')
        return
    if os.path.exists(consolidated_file):
        result = pd.read_excel(consolidated_file, sheet_name=None)
    else:
        result = {}
    entities = args['entities']
    for e in entities:
        if e in result:
            result[e] = result[e].fillna(0)
        else:
            result[e] = pd.DataFrame(columns=args['report_columns'], index=[])
        result[e]['ref_curr'] = result[e]['ref_curr'].astype(str)
        result[e]['ref_time'] = result[e]['ref_time'].astype(str)
        result[e]['reference'] = result[e]['ref_curr'] + result[e]['ref_time']
    refs = {e: set(result[e]['reference']) for e in entities}
    client_set = set.union(*[set(result[e]['client']) for e in entities]) - {0}

    to_add = {e: [] for e in entities}
    print('adding data from reports:')
    for r in new_files:
        try:
            report = pd.read_excel(r).fillna(0)
            new, entity, refs = iz_process_report(report, refs, client_set, entities, args)
            if entity == 'none':
                print(f'No entity detected in {r} - please check the file!')
            elif entity == 'new':
                print(f'New entity detected in {r}! add it to config file and process again')
            else:
                to_add[entity] += new
                print(f'{r} - success, {len(new)} lines')
        except Exception as ex:
            print(f'failed to process {r}')
            print(ex)
    for e in entities:
        if to_add[e]:
            df = pd.DataFrame(to_add[e])
            df['ref_curr'] = df['reference'].astype(str).apply(lambda v: v[:-17])
            df['ref_time'] = df['reference'].astype(str).apply(lambda v: v[-17:])
            result[e] = pd.concat([result[e], df], axis=0).sort_values('ref_time')
        result[e].drop('reference', axis=1, inplace=True)
    save_excel_multiple_sheets(result, consolidated_file)
    print(f'consolidated Izbank report refreshed at {datetime.now()}')


def vb_consolidated(args):
    target_directory = args["target_directory"]
    Path(target_directory).mkdir(parents=True, exist_ok=True)
    consolidated_file = args['consolidated_file']
    memory_file = args["memory_file"]
    try:
        with open(memory_file, "r", encoding='utf-8') as f:
            memory = json.load(f)
    except Exception:
        memory = []
    new_files = [v for v in os.listdir(target_directory) if (not v.startswith('.') and v not in memory)]
    if not new_files:
        print(f'No files to process in {target_directory}')
        return
    if os.path.exists(consolidated_file):
        result = pd.read_excel(consolidated_file, sheet_name=None)
    else:
        result = {}
    entities = args['entities']
    for e in entities:
        if e in result:
            result[e] = result[e].fillna(0)
        else:
            result[e] = pd.DataFrame(columns=args['report_columns'], index=[])
        result[e]['reference'] = result[e]['reference'].astype(str)
    refs = {e: set(result[e]['reference']) for e in entities}
    client_set = set.union(*[set(result[e]['client']) for e in entities]) - {0}

    to_add = {e: [] for e in entities}
    print('adding data from reports:')
    for r in new_files:
        try:
            report = pd.read_excel(r).fillna(0)
            new, entity, refs = vb_process_report(report, refs, client_set, entities, args)
            if entity == 'none':
                print(f'No entity detected in {r} - please check the file!')
            elif entity == 'new':
                print(f'New entity detected in {r}! add it to config file and process again')
            else:
                to_add[entity] += new
                print(f'{r} - success, {len(new)} lines')
        except Exception as ex:
            print(f'failed to process {r}')
            print(ex)
    for e in entities:
        if to_add[e]:
            df = pd.DataFrame(to_add[e])
            result[e] = pd.concat([result[e], df], axis=0).sort_values('reference')
    save_excel_multiple_sheets(result, consolidated_file)
    print(f'consolidated Vakifbank report refreshed at {datetime.now()}')


def extract_vb_files_from_mail(args):
    print('Process of copying VB reports started', str(datetime.now().date()), str(datetime.now().time())[:5])
    target_directory = args["target_directory"]
    memory_file = args["memory_file"]
    try:
        with open(memory_file, "r", encoding='utf-8') as f:
            memory = json.load(f)
    except Exception:
        memory = []
    Path(target_directory).mkdir(parents=True, exist_ok=True)
    inbox = outlook.GetDefaultFolder(6).folders(args["folder"])
    new_files = []
    for msg in inbox.Items:
        ats = msg.Attachments
        if len(ats) == 1:
            att = ats.Item(1)
            f_temp = os.path.join(working_directory, "temp.xlsx")
            att.SaveASFile(f_temp)
            try:
                df = pd.read_excel(f_temp)
                curr = df.iloc[2, 1].split()[1]
                date = str(parser.parse(df.columns[7], dayfirst=True))[:10]
                f_new = f'{str(att)[:-5]}.{curr}.{date}.xlsx'
                f_new_full_path = f'{target_directory}{str(att)[:-5]}.{curr}.{date}.xlsx'
                if f_new not in memory:
                    print(f'got new file with currency = {curr}, date = {date}, result file={f_new}')
                    shutil.copy(f_temp, f_new_full_path)
                    memory.append(f_new)
                    new_files.append(f_new_full_path)
            except Exception as ex:
                print(f'{ex} : file {str(att)}')
            os.remove(f_temp)
    with open(memory_file, "w", encoding='utf-8') as f:
        json.dump(memory, f)
    if new_files:
        print(f"VB Bank: {len(new_files)} files extracted")
    else:
        print("VB Bank: no files")


def extract_iz_files_from_mail(args):
    print('Process of copying IZ reports started', str(datetime.now().date()), str(datetime.now().time())[:5])
    target_directory = args["target_directory"]
    memory_file = args["memory_file"]
    try:
        with open(memory_file, "r", encoding='utf-8') as f:
            memory = json.load(f)
    except Exception:
        memory = []
    Path(target_directory).mkdir(parents=True, exist_ok=True)
    inbox = outlook.GetDefaultFolder(6).folders(args["folder"])
    new_files = []
    for msg in inbox.Items:
        ats = msg.Attachments
        if len(ats) == 1:
            att = ats.Item(1)
            f_name = str(att)
            if (f_name[-5:] != ".xlsx" and f_name[-4:] != ".xls") or f_name in memory:
                continue
            memory.append(f_name)
            print(f'got new file {f_name}')
            f_new_full_path = f'{target_directory}{f_name}'
            att.SaveASFile(f_new_full_path)
            new_files.append(f_new_full_path)
    with open(memory_file, "w", encoding='utf-8') as f:
        json.dump(memory, f)
    if new_files:
        print(f"IZ bank: {len(new_files)} files extracted, now processing")
    else:
        print("IZ bank: no files")


def main(config, pause=10):
    start = config['start_list']
    regulars = {}
    for v in start:
        if start[v][0].startswith('min'):
            regulars[v] = int(start[v][0][3:])

    # function gets next trigger time for a particular item in start list
    def get_next_trigger(item):
        today = str(datetime.now().date())
        time_now = str(datetime.now().time())[:5]
        if item in regulars:
            return datetime.now() + timedelta(minutes=regulars[item])
        left_today = [t for t in start[item] if t > time_now]
        if left_today:
            return parser.parse(f'{today} {min(left_today)}')
        else:
            tomorrow = str(datetime.now() + timedelta(days=1))[:10]
            return parser.parse(f'{tomorrow} {min(start[item])}')

    now = datetime.now()
    next_trigger = {v: now if v in regulars else get_next_trigger(v) for v in start}
    while True:
        now = datetime.now()
        for v in start:
            if now > next_trigger[v]:
                next_trigger[v] = get_next_trigger(v)
                args = (config.get(v, None),)
                Process(target=globals()[v], args=args).start()
        time.sleep(pause)
