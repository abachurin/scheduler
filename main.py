from base.scheduler import *


if __name__ == '__main__':

    try:
        import win32com.client as win32

        xl_app = win32.DispatchEx("Excel.Application")
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as ex:
        print("No Windows module")

    # get configuration
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)

    # run the scheduler
    print(1)
    main(config)
