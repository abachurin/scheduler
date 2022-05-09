from scheduler import *

name = 'extract_iz_files_from_mail'

with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)
args = config.get(name, None)
globals()[name](args)
