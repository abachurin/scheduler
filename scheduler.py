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


def add_report_to_consolidated(args):
    directory_list = args['directory']
    consolidated_file = args['consolidated_file']
    entities = args['entities']

    # get memory or create empty if doesn't exist yet
    all_files = []
    for d in directory_list:
        all_files += [os.path.join(d, v) for v in os.listdir(d) if not v.startswith('.')]

    if os.path.exists(consolidated_file):
        result = pd.read_excel(consolidated_file, sheet_name=None)
    else:
        result = {}
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
    for r in all_files:
        try:
            report = pd.read_excel(r).fillna(0)
            new, entity = process_report(report, refs, client_set, entities, args)
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
    print(f'consolidated report refreshed at {datetime.now()}')


def extract_vb_files_from_mail(args):
    time.sleep(5)
    print('Job 3. Process of copying VB reports started', str(datetime.now().date()), str(datetime.now().time())[:5])
    target_directory = args["target_directory"]
    memory_file = args["memory_file"]
    try:
        with open(memory_file, "r") as f:
            memory = json.load(f)
    except Exception:
        memory = []
    Path(target_directory).mkdir(parents=True, exist_ok=True)
    inbox = outlook.GetDefaultFolder(6).folders(args["folder"])
    counter = 0
    for msg in inbox.Items:
        ats = msg.Attachments
        if len(ats) == 1:
            att = ats.Item(1)
            f = str(att)
            f_temp = os.path.join(working_directory, "temp.xlsx")
            att.SaveASFile(f_temp)
            try:
                df = pd.read_excel(f_temp)
                curr = df.iloc[2, 1][-3:]
                date = str(parser.parse(df.columns[7], dayfirst=True))[:10]
                f_new = f'{target_directory}{f[:-5]}.{curr}.{date}.xlsx'
                if f_new not in memory:
                    print(f'Job 3. got new file with currency = {curr}, date = {date}, result file={f_new}')
                    shutil.copy(f_temp, f_new)
                    memory.append(f_new)
                    counter += 1
            except Exception as ex:
                print(ex)
            os.remove(f_temp)
    with open(memory_file, "w") as f:
        json.dump(memory, f)
    if counter:
        print(f"Job 3: {counter} files processed")
    else:
        print("Job 3: no files")


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


if __name__ == '__main__':

    # get configuration
    with open('config.json', 'r', encoding='utf8') as f:
        config = json.load(f)

    # run the scheduler
    main(config)
