from base.processing import *

path_1 = 'my_excel_file.xlsx'
path_2 = 'another_excel_file.xlsx'


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
    directory = args['directory']
    consolidated_file = args['consolidated_file']
    memory_file = args['memory']

    # get memory or create empty if doesn't exist yet
    all_files = [v for v in os.listdir(directory) if not v.startswith('.')]
    pprint(all_files)
    if memory_file not in all_files:
        content = {
            'skip_files': [consolidated_file, memory_file],
            'clients': []
        }
        with open(directory + memory_file, 'w') as f:
            json.dump(content, f)
    with open(directory + memory_file, 'r') as f:
        memory = json.load(f)

    if consolidated_file in all_files:
        result = pd.read_excel(directory + consolidated_file)
    else:
        result = pd.DataFrame(columns=['transaction_date'], index=[])
        pprint(result)

    files_to_add = [v for v in all_files if v not in memory['skip_files']]
    print(f'adding data from reports: {files_to_add}')
    client_list = memory['clients']
    for r in files_to_add:
        try:
            report = pd.read_excel(directory + r)
            memory['skip_files'].append(r)
        except Exception as ex:
            print(f'failed to process {r}')
            print(ex)

    result.to_excel(directory + consolidated_file, index=False)
    with open(directory + memory_file, 'w') as f:
        json.dump(memory, f)


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
