# using "__getattr__" without a class 使用getattr用字串指定function名稱
import sys
import schedule
import time
import threading
import configparser
import datetime
from win32com import ShellExecute as ShellExecute

# 當前文件
current_module = sys.modules[__name__]
config = configparser.ConfigParser()
config.read('AutoDownload_Setting.ini')


def run_application():
    app_path = config['setting']['app_path']

    # ShellExecute(父視窗, 動作, 檔案路徑, 參數(檔案的話是空值), 程式初始化目錄, 是否顯示視窗)
    ShellExecute(0, 'open', app_path, '', '', 1)


# def run_application2():
#     app_path2 = config['setting']['app_path2']
#
#     # ShellExecute(父視窗, 動作, 檔案路徑, 參數(檔案的話是空值), 程式初始化目錄, 是否顯示視窗)
#     ShellExecute(0, 'open', app_path2, '', '', 1)


def run_schedule():

    set_time = config['setting']['set_time']

    # function_name = config['setting']['function_name']

    # every day at specific time
    schedule.every().day.at(set_time).do(getattr(current_module, "run_application"))

    # set_time2 = config['setting']['set_time2']
    # function_name2 = config['setting']['function_name2']
    # schedule.every().day.at(set_time2).do(getattr(current_module, function_name2))

    flag = 0

    while True:
        current_time = datetime.datetime.now()

        if current_time.strftime('%S') == '00':

            if flag == 0 or int(current_time.strftime('%M')) % 5 == 0:
                print('現在時間: {} 執行中...'.format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                flag = 1

            schedule.run_pending()

        time.sleep(1)


thread = threading.Thread(target=run_schedule)
thread.start()
