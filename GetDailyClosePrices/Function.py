import datetime
import time
import requests
import pandas as pd
import os
from sqlalchemy import create_engine
import configparser
from exchange_calendars import get_calendar

config = configparser.ConfigParser()
config.read('GetDailyClosePrices_Setting.ini')
sql_path = config['sql']['sql_path']
SQLEngine = create_engine(sql_path)


def get_daily_close_prices():
    date = datetime.date.today().strftime('%Y%m%d')
    # date = '20220318'
    tw_calendar = get_calendar('XTAI')

    if date not in tw_calendar.opens:
        print('今日非交易日')
        time.sleep(3)
        return

    # 網址
    url = 'https://www.twse.com.tw/exchangeReport/MI_INDEX'

    # 參數
    user_agent = config['html']['user_agent']

    params = {
        'response': 'html',
        'date': date,
        'type': 'ALLBUT0999'
    }

    headers = {
        'user-agent': user_agent
    }
    response = requests.get(url, params=params, headers=headers)
    response = response.text.replace('div class="table-caption', 'tr class="table-caption')
    response = response.replace('div class="table-row', 'tr class="table-row')
    response = response.replace('span class="table-cell', 'td class="table-cell')
    try:
        # 取最後一個表格的html存入dataframe
        df = pd.read_html(response)[-1]
    except:
        return None

    # 欄位名稱刪除columns(1)與columns(0)，只留columns(2)
    # df.columns = df.columns.droplevel(1).droplevel(0)

    # 直接設定columns(2)為df欄位名稱
    df.columns = df.columns.get_level_values(2)

    # axis=0代表根據索引值刪除列，axis=1代表根據欄位名稱刪除行
    df.drop(['證券名稱', '漲跌(+/-)'], axis=1, inplace=True)

    # 將date轉換成datetime後加入df日期欄位
    df['日期'] = pd.to_datetime(date)

    # 設定index為證券代號與日期
    df.set_index(['證券代號', '日期'], inplace=True)

    # 將df中美個欄位的資料轉為數字，無法轉換則errors='coerce' 表示NaN
    df = df.apply(pd.to_numeric, errors='coerce')

    # 顯示收盤價為NaN的資料列
    df[df['收盤價'].isnull()]

    # 刪除收盤價為NaN的資料列
    df.drop(df[df['收盤價'].isnull()].index, inplace=True)

    excel_name = '單日收盤行情({}).xlsx'.format(date)
    sheet_name = '{}'.format(date)
    # to_excel(excel_name,sheet_name)
    df.to_excel(excel_name, sheet_name)

    # 檢查檔案是否建立成功，成功後新增置資料庫juridical_person_buy_over
    if os.path.isfile(excel_name):

        print(f'{excel_name} 檔案建立完成')
        try:
            df.to_sql('daily_close_prices', SQLEngine, if_exists='append')

            print('現在時間: {} 寫入資料庫成功...'.format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))

        except:
            print('現在時間: {} 寫入資料庫失敗...'.format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    else:
        print('現在時間: {} {} 檔案建立失敗'.format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), excel_name))

    time.sleep(5)
    return df


get_daily_close_prices()
