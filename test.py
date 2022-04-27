from scheduler import *

name = 'add_report_to_consolidated'

with open('config.json', 'r') as f:
    config = json.load(f)
args = config.get(name, None)
globals()[name](args)
