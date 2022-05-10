from base.scheduler import *


if __name__ == '__main__':

    # get configuration
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)

    # run the scheduler
    print(1)
    main(config)
