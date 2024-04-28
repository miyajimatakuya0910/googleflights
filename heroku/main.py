from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.service import Service as ChromeService
from concurrent.futures import ThreadPoolExecutor
import time
from time import sleep
import re
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import csv
import os
from slack_sdk import WebClient
import requests
import traceback
import psutil
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from gspread.exceptions import WorksheetNotFound, SpreadsheetNotFound

# 指定したいフォルダのIDを設定 # 区間毎に変更
folder_id = '###########'
chromedriver_path = str(os.environ.get('CHROMEDRIVER_PATH'))

# Configure WebDriver options
options = webdriver.ChromeOptions()
options.add_argument('--headless')  # Run Chrome in headless mode (no GUI)
options.add_argument('--disable-gpu')  # Disable GPU acceleration (useful for headless mode)
options.add_argument('--no-sandbox')
options.add_argument("--disable-dev-shm-usage")

# ドライバーの起動
driver = webdriver.Chrome(executable_path=chromedriver_path, options=options)
driver.implicitly_wait(10)
wait = WebDriverWait(driver, 10)
# ウインドウサイズ拡大
driver.set_window_size('1200', '1000')

# 現在時刻（処理開始前）を取得
start = time.time()
print(start)

# CSVファイルからURLリストを読み込む
url_df = pd.read_csv('urls.csv', encoding="utf-8")
urls = url_df['URL'].tolist()
routes = url_df['route'].tolist()

# 今日の日付を取得
today = datetime.today()
today1 = datetime.today().strftime('%Y-%m-%d')

# 62日前の日付を計算し、逆順にリストに格納
datelist = []
for i in range(63):  # 0から62までの範囲を取得
    date = today - timedelta(days=i)
    datelist.insert(0, date.strftime('%Y-%m-%d'))

# x日後の日付を計算し、リストに格納
datelist1 = []
for i in range(328): # リストの日付数に対応
    date = today + timedelta(days=i)
    datelist1.append(date.strftime('%Y-%m-%d'))

# 数値だけを取り出す関数
def extract_numbers(text):
    # 文字列内の数字以外の文字を削除し、'$' も削除する
    numeric_text = ''.join(filter(str.isdigit, text))
    
    if numeric_text:
        # 数値文字列を整数または浮動小数点数に変換
        numeric_value = int(numeric_text)  # 整数に変換する場合
    else:
        numeric_value = text

    return numeric_value

# データ取得を関数化
def getinfo():
    # Navigate to the Google Flights URL
    # driver.get(url)
    # wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.Xsgmwe.sI2Nye')))
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.n9rd7b.ZSxxwc')))
    tab = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-k8QpJ.VfPpkd-LgbsSe-OWXEXe-Bz112c-M1Soyc.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.nCP5yc.AjY5Oe.LQeN7.nJawce.OTelKf.XPGpHc.mAozAc')))
    tab.click()

    # ページのHTMLソースコードを取得
    html_source = driver.page_source

    # # 正規表現を使用してフライト名を抽出
    # flight_name_pattern = re.compile(r"Flight number (\w+ \d+)")
    # matches = flight_name_pattern.search(html_source)

    # if matches:
    #     flight_name = matches.group(1)
    # else:
    #     flight_name = "フライト名が見つかりません"
    
    time.sleep(1)
    try:
        flight_name = driver.find_element_by_xpath("/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div/div[2]/div[2]/div[3]/div/div[2]/div[1]/div/div/div/div/div[2]/div/div[5]/div[1]/div[9]/span[10]").text
        print(flight_name)
    except NoSuchElementException:
        print("No flightname")
        message = routes[0] + " Classname might have been changed"
        send_message2(message)

    df.loc[flight_name] = np.nan
    for d in range(4):
        df.loc[f'{flight_name}Price {d}'] = np.nan
        df.loc[f'{flight_name}Cabin Name {d}'] = np.nan
    
    # 出発時刻
    dep = driver.find_element_by_css_selector('.b0EVyb.YMlIz.y52p7d').get_attribute("textContent")
    df.iat[-9, 0] = dep
    # 到着時刻
    arr = driver.find_element_by_css_selector('.OJg28c.YMlIz.y52p7d').get_attribute("textContent")
    df.iat[-9, 1] = arr

    wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'Xsgmwe')))
    try:
        flightinfo = driver.find_elements_by_class_name('Xsgmwe')
        # 航空会社
        for airline in flightinfo:
            time.sleep(0.5)
            df.iat[-9, 2] = airline.text
            break
        # 運行機材
        for fleet in flightinfo[3:]:
            df.iat[-9, 3] = fleet.text
            break
    except NoSuchElementException:
        print("No flightinfo")

    # グラフの価格を入手
    index = df.index.get_loc(flight_name)
    k = 4
    if re.search(r'Price history', html_source):
        for element in driver.find_elements_by_css_selector('.ke9kZe-LkdAo-RbRzK-JNdkSc.pKrx3d'):
            text = element.get_attribute("aria-label")
            # 正規表現を使って金額を抽出
            match = re.search(r'\$([\d,]+)', text)
            if match:
                amount_str = match.group(1)
                # コンマを削除して数値に変換
                amount = int(amount_str.replace(',', ''))
                # DataFrameに金額を追加
                df.iat[index, k] = amount
                k += 1

        if "Today" in text:
            # print( 'There is "Today" ' )
            # 欠損値をチェックし、条件が満たされた場合に行を右にシフトする
            while pd.isna(df.iloc[index, -1]):
                df.iloc[index, 4:] = df.iloc[index, 4:].shift(1)
        else:
            # 欠損値をチェックし、条件が満たされた場合に行を右にシフトする
            # print( 'There is no "Today" ' )
            while pd.isna(df.iloc[index, -2]):
                df.iloc[index, 4:] = df.iloc[index, 4:].shift(1)
            try:
                # 今日の価格
                price_today_element = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div/div[2]/div[2]/div[2]/div/div/div[2]/div/div[1]/span')
                price_today = price_today_element.text
                df.loc[flight_name, df.columns[-1]] = extract_numbers(price_today)
            except NoSuchElementException:
                print(c)
                print("Price unavailable")
    else:
        try:
            # 今日の価格
            price_today_element = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div/div[2]/div[2]/div[2]/div/div/div[2]/div/div[1]/span')
            price_today = price_today_element.text
            df.loc[flight_name, df.columns[-1]] = extract_numbers(price_today)
            print(price_today)
        except NoSuchElementException:
            print(c)
            print("Price unavailable")

    # 価格,Cabin情報を取得
    try:
        price = driver.find_elements_by_class_name('tZe0ff')
        cabin = driver.find_elements_by_css_selector('.DllrY.YMlIz.ogfYpf')
        a = 8
        b = 0
        for l in price[0:4]:
            df.iat[-a, -1] = extract_numbers(l.text)
            df.iat[-a+1, -1] = cabin[b].text
            a -= 2
            b += 1
    except NoSuchElementException:
        print(c)
        print("No cabininfo")
        # message = routes[0] + " Classname might be changed"
        # send_message2(message)

# 日付アクセス
def dateaccess(f):
    textbox = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR,'.TP4Lpb.eoY5cb.j0Ppje')))
    time.sleep(0.5)
    textbox[0].click()
    time.sleep(0.2)
    textbox[2].clear()
    textbox[2].send_keys(datelist1[f])
    done = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-k8QpJ.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.nCP5yc.AjY5Oe.DuMIQc.LQeN7.z18xM.rtW97.Q74FEc.dAwNDc')))
    ActionChains(driver).move_to_element(done).perform()
    done.click()
    wait.until(EC.presence_of_all_elements_located)

# Slackメッセージの通知(#general)
def send_message(message):
    url = "https://slack.com/api/chat.postMessage"
    data = {
        "token": "#############",
        "channel": "C05SGFGT0D7",
        "text": message
    }
    requests.post(url, data=data)

# Slackメッセージの通知(#googleflights)
def send_message2(message):
    url = "https://slack.com/api/chat.postMessage"
    data = {
        "token": "#############",
        "channel": "C05S3S9NJSE",
        "text": message
    }
    requests.post(url, data=data)

# 2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
scope = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive']
#ダウンロードしたjsonファイル名をクレデンシャル変数に設定。
credentials = Credentials.from_service_account_file("iconic-medium-402306-d13ecc69d6aa.json", scopes=scope)
#OAuth2の資格情報を使用してGoogle APIにログイン。
gc = gspread.authorize(credentials)

driver.get(urls[0])
f = 0#変更
for day in datelist1:#変更
    dateaccess(f)
    print(datelist1[f])
    time.sleep(2)
    cur_url = driver.current_url
    #ファイルの有無確認
    try:
        ss = gc.open(datelist1[f], folder_id=folder_id)
        st = ss.get_worksheet(0)
        df = get_as_dataframe(st, index_col = 0, header = 0).dropna(how='all')
        # if not today1 in df.columns:
        df[today1] = np.nan
        c = 0 #0
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.zBTtmb.ZSxxwc')))
        try:
            buttons = driver.find_elements_by_css_selector('.VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-k8QpJ.VfPpkd-LgbsSe-OWXEXe-Bz112c-M1Soyc.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.nCP5yc.AjY5Oe.LQeN7.nJawce.OTelKf.XPGpHc.mAozAc')
        except NoSuchElementException:
            buttons = []
        if len(buttons) != 0:
            try:
                top = driver.find_element_by_class_name('VfPpkd-WsjYwc-OWXEXe-INsAgc')
            except NoSuchElementException:
                top = None 
            # 上のタブが存在するとき
            if top is not None:
                # ボタンを全て押下
                length = len(buttons)-1
            # 上のタブが存在しないとき
            else:
                length = len(buttons)

            for x in range(length):#1
                try:
                    # ボタン押下
                    buttons = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-k8QpJ.VfPpkd-LgbsSe-OWXEXe-Bz112c-M1Soyc.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.nCP5yc.AjY5Oe.LQeN7.nJawce.OTelKf.XPGpHc.mAozAc')))
                    # 上のタブが存在するとき
                    if length != len(buttons): 
                        for button in buttons[c+1:]:
                            ActionChains(driver).move_to_element(button).perform()
                            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-k8QpJ.VfPpkd-LgbsSe-OWXEXe-Bz112c-M1Soyc.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.nCP5yc.AjY5Oe.LQeN7.nJawce.OTelKf.XPGpHc.mAozAc')))
                            button.click()
                            wait.until(EC.presence_of_all_elements_located)
                            break
                        # ボタン押下
                        selectflights = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-INsAgc.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.Rj2Mlf.OLiIxf.PDpWxe.P62QJc.LQeN7.my6Xrf.wJjnG.dA7Fcf')))
                        for selectflight in selectflights[c:]:
                            ActionChains(driver).move_to_element(selectflight).perform()
                            selectflight.click()
                            break                     
                    else: 
                        for button in buttons[c:]:
                            ActionChains(driver).move_to_element(button).perform()
                            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-k8QpJ.VfPpkd-LgbsSe-OWXEXe-Bz112c-M1Soyc.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.nCP5yc.AjY5Oe.LQeN7.nJawce.OTelKf.XPGpHc.mAozAc')))
                            button.click()
                            wait.until(EC.presence_of_all_elements_located)
                            break
                        # ボタン押下
                        selectflights = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-INsAgc.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.Rj2Mlf.OLiIxf.PDpWxe.P62QJc.LQeN7.my6Xrf.wJjnG.dA7Fcf')))
                        for selectflight in selectflights[c:]:
                            ActionChains(driver).move_to_element(selectflight).perform()
                            selectflight.click()
                            break

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.n9rd7b.ZSxxwc')))
                    tab = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-k8QpJ.VfPpkd-LgbsSe-OWXEXe-Bz112c-M1Soyc.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.nCP5yc.AjY5Oe.LQeN7.nJawce.OTelKf.XPGpHc.mAozAc')))
                    tab.click()

                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.n9rd7b.ZSxxwc')))

                    # ページのHTMLソースコードを取得
                    html_source = driver.page_source

                    # # 正規表現を使用してフライト名を抽出
                    # flight_name_pattern = re.compile(r"Flight number (\w+ \d+)")
                    # matches = flight_name_pattern.search(html_source)

                    # if matches:
                    #     flight_name = matches.group(1)
                    # else:
                    #     flight_name = "フライト名が見つかりません"

                    time.sleep(1)
                    try:
                        flight_name = driver.find_element_by_xpath("/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div/div[2]/div[2]/div[3]/div/div[2]/div[1]/div/div/div/div/div[2]/div/div[5]/div[1]/div[9]/span[10]").text
                        print(flight_name)
                    except NoSuchElementException:
                        print("No flightname")
                        message = routes[0] + " Classname might have been changed"
                        send_message2(message)

                    # 行名が存在するか確認
                    if flight_name in df.index:
                        try:
                            price_today = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div/div[2]/div[2]/div[2]/div/div/div[2]/div/div[1]/span').text
                            # price_today = driver.find_element_by_class_name('QORQHb').text
                            df.loc[flight_name, today1] = extract_numbers(price_today)
                            print(price_today)
                        except NoSuchElementException:
                            print(c)
                            print("Price unavailable")

                        try:
                            price = driver.find_elements_by_class_name('tZe0ff')
                            cabin = driver.find_elements_by_css_selector('.DllrY.YMlIz.ogfYpf')
                            i = 0
                            for l in price[0:4]:
                                df.loc[f'{flight_name}Price {i}', today1] = extract_numbers(l.text)
                                df.loc[f'{flight_name}Cabin Name {i}', today1] = cabin[i].text
                                i += 1
                            print("Cabininfo exists")
                        except NoSuchElementException:
                            print(c)
                            print("No cabininfo")
                            message = routes[0] + " No cabininfo"
                            send_message(message)

                        # print(df)
                        print(datelist1[f])
                        print(f)
                        print(c)
                        driver.get(cur_url)
                        c += 1
                    else:
                    # フライトが追加された場合、過去の価格、フライト情報、過去データを取得
                        df.loc[flight_name] = np.nan
                        df.loc[f'{flight_name}Price 0'] = np.nan
                        df.loc[f'{flight_name}Cabin Name 0'] = np.nan
                        df.loc[f'{flight_name}Price 1'] = np.nan
                        df.loc[f'{flight_name}Cabin Name 1'] = np.nan
                        df.loc[f'{flight_name}Price 2'] = np.nan
                        df.loc[f'{flight_name}Cabin Name 2'] = np.nan
                        df.loc[f'{flight_name}Price 3'] = np.nan
                        df.loc[f'{flight_name}Cabin Name 3'] = np.nan
                        index = df.index.get_loc(flight_name)

                        # 出発時刻
                        dep = driver.find_element_by_css_selector('.b0EVyb.YMlIz.y52p7d').get_attribute("textContent")
                        df.loc[flight_name, 'Departure time'] = dep
                        # 到着時刻
                        arr = driver.find_element_by_css_selector('.OJg28c.YMlIz.y52p7d').get_attribute("textContent")
                        df.loc[flight_name, 'Arrival time'] = arr

                        wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'Xsgmwe')))
                        try:
                            flightinfo = driver.find_elements_by_class_name('Xsgmwe')
                            # 航空会社
                            for airline in flightinfo:
                                time.sleep(0.5)
                                df.loc[flight_name, 'Airline'] = airline.text
                                break
                            # 運行機材
                            for fleet in flightinfo[3:]:
                                df.loc[flight_name, 'Fleet'] = fleet.text
                                break
                        except NoSuchElementException:
                            print("No flightinfo")   

                        try:
                            price = driver.find_elements_by_class_name('tZe0ff')
                            cabin = driver.find_elements_by_css_selector('.DllrY.YMlIz.ogfYpf')
                            i = 0
                            for l in price[0:4]:
                                df.loc[f'{flight_name}Price {i}', today1] = extract_numbers(l.text)
                                df.loc[f'{flight_name}Cabin Name {i}', today1] = cabin[i].text
                                i += 1
                        except NoSuchElementException:
                            print(c)
                            print("No cabininfo")
                            message = routes[0] + " No cabininfo"
                            send_message(message)

                        # グラフの価格を入手
                        k = 4
                        if re.search(r'Price history', html_source):
                            for element in driver.find_elements_by_css_selector('.ke9kZe-LkdAo-RbRzK-JNdkSc.pKrx3d'):
                                text = element.get_attribute("aria-label")
                                # 正規表現を使って金額を抽出
                                match = re.search(r'\$([\d,]+)', text)
                                if match:
                                    amount_str = match.group(1)
                                    # コンマを削除して数値に変換
                                    amount = int(amount_str.replace(',', ''))
                                    # DataFrameに金額を追加
                                    df.iat[index, k] = amount
                                    k += 1

                            if "Today" in text:
                                # print( 'There is "Today" ' )
                                # 欠損値をチェックし、条件が満たされた場合に行を右にシフトする
                                while pd.isna(df.iloc[index, -1]):
                                    df.iloc[index, 4:] = df.iloc[index, 4:].shift(1)
                            else:
                                # 欠損値をチェックし、条件が満たされた場合に行を右にシフトする
                                # print( 'There is no "Today" ' )
                                while pd.isna(df.iloc[index, -2]):
                                    df.iloc[index, 4:] = df.iloc[index, 4:].shift(1)
                                try:
                                    # 今日の価格
                                    price_today_element = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div/div[2]/div[2]/div[2]/div/div/div[2]/div/div[1]/span')
                                    price_today = price_today_element.text
                                    df.loc[flight_name, df.columns[-1]] = extract_numbers(price_today)
                                except NoSuchElementException:
                                    print(c)
                                    print("Price unavailable")
                        else:
                            try:
                                # 今日の価格
                                price_today_element = driver.find_element_by_xpath('/html/body/c-wiz[2]/div/div[2]/c-wiz/div[1]/c-wiz/div/div[2]/div[2]/div[2]/div/div/div[2]/div/div[1]/span')
                                price_today = price_today_element.text
                                df.loc[flight_name, df.columns[-1]] = extract_numbers(price_today)
                            except NoSuchElementException:
                                print(c)
                                print("Price unavailable")                            

                        # print(df)
                        print(datelist1[f])
                        print(f)
                        print(c)
                        driver.get(cur_url)
                        c += 1
                except TimeoutException as e:
                    print("TimeoutException")
                    message = routes[0] + datelist1[f] + str(c) + "TimeoutException"
                    send_message(message)
                    driver.get(cur_url)
                    c += 1
                except Exception as e:
                    print(traceback.format_exc())
                    message = routes[0] + datelist1[f] + str(c) + traceback.format_exc()
                    send_message(message)
                    driver.get(cur_url)
                    c += 1

            # スプレッドシートにDataFrameを書き込む
            st.clear()
            set_with_dataframe(st, df.reset_index())
            print(datelist1[f])
            print("Sent to SpreadSheet")
            message = routes[0] + datelist1[f] + ' ' + str(f)
            send_message(message)
            f += 1
            driver.quit()
            time.sleep(1)
            # ドライバーの再起動
            driver = webdriver.Chrome(executable_path=chromedriver_path, options=options)
            driver.implicitly_wait(10)
            wait = WebDriverWait(driver, 10)
            driver.set_window_size('1200', '1000')
            driver.get(cur_url)
            print("Restarted")
        else:
            print(datelist1[f])
            print("No flights")
            f += 1

        # else:
        #     print(datelist1[f])
        #     print(f)
        #     print('Already Exists')
        #     f += 1

    # 追加分
    # スプレッドシートが見つからなかった場合の処理
    except SpreadsheetNotFound:
        df = pd.DataFrame(columns=datelist)
        df.insert(loc = 0, column= 'Departure time', value= np.nan)
        df.insert(loc = 1, column= 'Arrival time', value= np.nan)
        df.insert(loc = 2, column= 'Airline', value= '')
        df.insert(loc = 3, column= 'Fleet', value= '')
        c = 0 #0
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.zBTtmb.ZSxxwc')))
        try:
            buttons = driver.find_elements_by_css_selector('.VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-k8QpJ.VfPpkd-LgbsSe-OWXEXe-Bz112c-M1Soyc.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.nCP5yc.AjY5Oe.LQeN7.nJawce.OTelKf.XPGpHc.mAozAc')
        except NoSuchElementException:
            buttons = []
        if len(buttons) != 0:
            try:
                top = driver.find_element_by_class_name('VfPpkd-WsjYwc-OWXEXe-INsAgc')
            except NoSuchElementException:
                top = None 
            # 上のタブが存在するとき
            if top is not None:
                # ボタンを全て押下
                length = len(buttons)-1
            # 上のタブが存在しないとき
            else:
                length = len(buttons)

            for x in range(length):#1
                buttons = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-k8QpJ.VfPpkd-LgbsSe-OWXEXe-Bz112c-M1Soyc.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.nCP5yc.AjY5Oe.LQeN7.nJawce.OTelKf.XPGpHc.mAozAc')))
                # 上のタブが存在するとき
                if length != len(buttons): 
                    # ボタン押下
                    for button in buttons[c+1:]:
                        ActionChains(driver).move_to_element(button).perform()
                        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-k8QpJ.VfPpkd-LgbsSe-OWXEXe-Bz112c-M1Soyc.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.nCP5yc.AjY5Oe.LQeN7.nJawce.OTelKf.XPGpHc.mAozAc')))
                        button.click()
                        wait.until(EC.presence_of_all_elements_located)
                        break

                    # ボタン押下
                    selectflights = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-INsAgc.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.Rj2Mlf.OLiIxf.PDpWxe.P62QJc.LQeN7.my6Xrf.wJjnG.dA7Fcf')))
                    for selectflight in selectflights[c:]:
                        ActionChains(driver).move_to_element(selectflight).perform()
                        selectflight.click()
                        break
                else:
                    # ボタン押下
                    for button in buttons[c:]:
                        ActionChains(driver).move_to_element(button).perform()
                        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-k8QpJ.VfPpkd-LgbsSe-OWXEXe-Bz112c-M1Soyc.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.nCP5yc.AjY5Oe.LQeN7.nJawce.OTelKf.XPGpHc.mAozAc')))
                        button.click()
                        wait.until(EC.presence_of_all_elements_located)
                        break
                    # ボタン押下
                    selectflights = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.VfPpkd-LgbsSe.VfPpkd-LgbsSe-OWXEXe-INsAgc.VfPpkd-LgbsSe-OWXEXe-dgl2Hf.Rj2Mlf.OLiIxf.PDpWxe.P62QJc.LQeN7.my6Xrf.wJjnG.dA7Fcf')))
                    for selectflight in selectflights[c:]:
                        ActionChains(driver).move_to_element(selectflight).perform()
                        selectflight.click()
                        break
                try:
                    getinfo()
                except TimeoutException as e:
                    print("TimeoutException")
                    message = routes[0] + datelist1[f] + str(c) + "TimeoutException"
                    send_message(message)
                except Exception as e:
                    print(traceback.format_exc())
                    message = routes[0] + datelist1[f] + str(c) + traceback.format_exc()
                    send_message(message)
                # print(df)
                print(datelist1[f])
                print(f)
                print(c)
                driver.get(cur_url)
                c += 1
            # スプレッドシートにDataFrameを書き込む 
            ss = gc.create(datelist1[f], folder_id=folder_id)
            worksheet = ss.get_worksheet(0)
            set_with_dataframe(worksheet, df.reset_index())
            print(datelist1[f])
            print("Sent to SpreadSheet")
            message = routes[0] + datelist1[f] + ' ' + str(f)
            send_message(message)
            f += 1
            driver.quit()
            time.sleep(1)
            # ドライバーの再起動
            driver = webdriver.Chrome(executable_path=chromedriver_path, options=options)
            driver.implicitly_wait(10)
            wait = WebDriverWait(driver, 10)
            driver.set_window_size('1200', '1000')
            driver.get(cur_url)
            print("Restarted")
        else:
            print(datelist1[f])
            print("No flights")
            f += 1

    except Exception as e:
        # その他の例外が発生した場合の処理
        print(f'{routes[0]} スプレッドシート "{datelist1[f]}" を開く際にエラーが発生しました: {str(e)}')
        message = f'{routes[0]} スプレッドシート "{datelist1[f]}" を開く際にエラーが発生しました: {str(e)}'
        send_message(message)
        time.sleep(60)
        f += 1

# 現在時刻（処理完了後）を取得
end = time.time()
# 処理完了後の時刻から処理開始前の時刻を減算する
time_diff = (end - start)/3600
message = routes[0] + str(time_diff) + 'Completed'
send_message(message)
# WebDriverをクローズ
driver.quit()