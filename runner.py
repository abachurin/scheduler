from base.scheduler import *
import argparse
import sys


if __name__ == '__main__':

    # get configuration
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
    all_functions = list(config['start_list'])
    all_options = set([v for w in [set(config[func].get('commands', [])) for func in config['start_list']] for v in w])

    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument('-h', '--help', action="store_true")
    parser.add_argument('-n', '--no_saving', action="store_true")
    parser.add_argument('--func', help='choose function to run', choices=config['start_list'], default='none')
    for opt in all_options:
        parser.add_argument(f'--{opt}')

    args = parser.parse_args()
    name = args.func

    # Process help request
    if args.help:
        if name == 'none':
            print(config['help_message'])
            print(f'Available functions to start:\n{all_functions}')
        else:
            print(config[name].get('description', 'No description is available'))
            command_params = config[name].get('commands', [])
            if command_params:
                print('Available command parameters:\n'
                      '/add "-n" flag if you do NOT want to overwrite default values in config file/')
                for v in command_params:
                    print(f'--{v} : {config[name]["commands"][v]}')
        sys.exit()

    # collect arguments, change config file if necessary, and run chosen function
    print(f'running {name}\n-------------')
    func_args = config.get(name, {}).get('args', {})
    if func_args:
        for v in func_args:
            if v in args:
                func_args[v] = getattr(args, v)
        if not args.no_saving:
            config[name]['args'] = func_args
            with open('config.json', 'w', encoding='utf-8') as f:
                json.dump(config, f)
    globals()[name](func_args)
    print(f'{name} run is over\n-------------')

