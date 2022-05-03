from scheduler import *

name = 'extract_vb_files_from_mail'

with open('config.json', 'r') as f:
    config = json.load(f)
args = config.get(name, None)
globals()[name](args)
