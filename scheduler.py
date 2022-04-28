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
    with open('config.json', 'r') as f:
        config = json.load(f)

    # if 'excel' is ticked on in config file - try to run MS Excel. Works only under Windows OS
    if config['excel']:
        try:
            import win32com.client as win32
            xl_app = win32.DispatchEx("Excel.Application")
            print('MS Excel copy running')
        except Exception as ex:
            print(ex)
            print('MS Excel failed to start!')
    else:
        xl_app = None

    # run the scheduler
    main(config)
