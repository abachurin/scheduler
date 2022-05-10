from base.scheduler import *

try:
    import win32com.client as win32
    xl_app = win32.DispatchEx("Excel.Application")
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
except Exception as ex:
    print("No Windows module")

name = 'extract_vb_files_from_mail'

with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)
args = config.get(name, None)
globals()[name](args)
