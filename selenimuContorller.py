from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import time
import random

import pytz
tokyo_tz = pytz.timezone('Asia/Tokyo')

class selenimuContorller():
    driver = ""
    actionCount = 0

    INTERVAL_TIME=3 # リスクエス間隔は3秒に仮置き
    last_processed_time = 0

    def __init__(self, isCloud=False):
        self.actionlog(f'[open] ブラウザを開きます。')
        # オプション指定
        options = webdriver.ChromeOptions()
        if isCloud:
            options.add_argument('--headless') # Cloud化のためheadlessオプションを有効化
        options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36')
        options.add_argument('--disable-blink-features=AutomationControlled')
        self.driver = webdriver.Chrome(options=options)

    def getUrl(self, url):
        self.actionlog(f'[get] [ {url} ]を表示します。')
        self.driver.get(url)

    def close(self):
        self.actionlog(f'[close] ブラウザを閉じます。')
        self.driver.close()
        self.driver.quit()

    def wait(self, waitTime, elementType, elementValue):
        self.wait_any_of(waitTime, elementType, elementValue)

    def wait_any_of(self, waitTime, elementType, elementValue,
                                    elementType2='', elementValue2='',
                                    elementType3='', elementValue3='',
                                    elementType4='', elementValue4=''):
        if elementType2 == '':
            self.actionlog(f'[wait] 要素[{elementType},{elementValue}]の表示を最大[{waitTime}]秒待ちます。')
        elif elementType3 == '':
            self.actionlog(f'[wait] 要素[{elementType},{elementValue}]、もしくは、要素[{elementType2},{elementValue2}]の表示を最大[{waitTime}]秒待ちます。')
        elif elementType4 == '':
            self.actionlog(f'[wait] 要素[{elementType},{elementValue}]、もしくは、要素[{elementType2},{elementValue2}]、もしくは、要素[{elementType3},{elementValue3}]の表示を最大[{waitTime}]秒待ちます。')
        else:
            self.actionlog(f'[wait] 要素[{elementType},{elementValue}]、もしくは、要素[{elementType2},{elementValue2}]、もしくは、要素[{elementType3},{elementValue3}]、もしくは、要素[{elementType4},{elementValue4}]の表示を最大[{waitTime}]秒待ちます。')
        
        # try:
        wait = WebDriverWait(self.driver, waitTime)
        if elementType2 == '':
            element = wait.until(EC.visibility_of_element_located((elementType, elementValue)))    
        elif elementType3 == '':
            wait.until(
                EC.any_of(
                    EC.visibility_of_element_located((elementType, elementValue)),
                    EC.visibility_of_element_located((elementType2, elementValue2))
                )
            )
        elif elementType4 == '':
            wait.until(
                EC.any_of(
                    EC.visibility_of_element_located((elementType, elementValue)),
                    EC.visibility_of_element_located((elementType2, elementValue2)),
                    EC.visibility_of_element_located((elementType3, elementValue3))
                )
            )
        else:
            wait.until(
                EC.any_of(
                    EC.visibility_of_element_located((elementType, elementValue)),
                    EC.visibility_of_element_located((elementType2, elementValue2)),
                    EC.visibility_of_element_located((elementType3, elementValue3)),
                    EC.visibility_of_element_located((elementType4, elementValue4))
                )
            )

        # except Exception as e:
        #     self.errorlog(f'エラーが発生しました: {e}')

        # 前回の処理時間からの経過時間
        elapsed_time = time.time() - self.last_processed_time
        self.log(f'経過時間：{elapsed_time:.3f}秒')
        # 経過時間がリクエスト間隔よりも短い場合にリクエストを待ち合わせる
        if elapsed_time < self.INTERVAL_TIME:
            sleep_time = (self.INTERVAL_TIME - elapsed_time) * random.uniform(0.8, 1.2)
            self.log(f'待機時間：{sleep_time:.3f}秒')
            time.sleep(sleep_time)
        self.last_processed_time = time.time()

        return

    def click(self, elementType, elementValue):
        self.actionlog(f'[click] 要素[{elementType},{elementValue}]をクリックします。')

        # if len(self.driver.find_elements(elementType, elementValue)) == 0:
        #     self.errorlog(f'エラーが発生しました: 要素[{elementType},{elementValue}]はありません。')
        #     return
        
        # try:
        self.driver.find_element(elementType, elementValue).click()
        # except Exception as e:
        #     self.errorlog(f'エラーが発生しました: {e}')

    def send_keys(self, elementType, elementValue, sendValue, hide_input_value=False):
        input_value = sendValue
        if hide_input_value:
            input_value = '**********'
        self.actionlog(f'[send_keys] 要素[{elementType},{elementValue}]に[{input_value}]を入力します。')
        
        # if len(self.driver.find_elements(elementType, elementValue)) == 0:
        #     self.errorlog(f"エラーが発生しました: 要素[{elementType},{elementValue}]はありません。")
        #     return

        # try:
        self.driver.find_element(elementType, elementValue).clear()
        self.driver.find_element(elementType, elementValue).send_keys(sendValue)

        # except Exception as e:
        #     self.errorlog(f'エラーが発生しました: {e}')

    def select(self, elementType, elementValue, selectValue):
        self.actionlog(f'[select] 要素[{elementType},{elementValue}]の[{selectValue}]を選択します。')

        # try:
        Select(self.driver.find_element(elementType, elementValue)).select_by_value(selectValue)

        # except Exception as e:
        #     self.errorlog(f'エラーが発生しました: {e}')

    def is_enabled(self, elementType, elementValue):
        if len(self.driver.find_elements(elementType, elementValue)) == 0:
            self.log(f'[check] 要素[{elementType},{elementValue}]は使用できません。')
            return False
        
        if self.driver.find_element(elementType, elementValue).is_enabled() == False:
            self.log(f'[check] 要素[{elementType},{elementValue}]は使用できません。')
            return False
        
        return True

    def focusToElement(self, elementType, elementValue):
        JavaScriptFocusToElement = "arguments[0].focus({'preventScroll': arguments[1]})"
        element = self.driver.find_element(elementType, elementValue)
        self.driver.execute_script(JavaScriptFocusToElement, element, True)

    def get_text(self, elementType, elementValue, i=0):
        return self.driver.find_elements(elementType, elementValue)[i].text
    
    def get_element_count(self, elementType, elementValue):
        return len(self.driver.find_elements(elementType, elementValue))

    def actionlog(self, logdata):
        self.actionCount += 1
        print(f'{datetime.now(tokyo_tz).strftime("%Y-%m-%d_%H:%M:%S.%f")} #{self.actionCount} {logdata}')

    def errorlog(self, logdata):
        print(f'{datetime.now(tokyo_tz).strftime("%Y-%m-%d_%H:%M:%S.%f")} {logdata}')

    def log(self, logdata):
        print(f'{datetime.now(tokyo_tz).strftime("%Y-%m-%d_%H:%M:%S.%f")} {logdata}')