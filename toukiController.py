## 不動産請求画面制御
import selenimuContorller
import xlsContorller
import gsContorller
import time
from datetime import datetime
import Message

MAX_CHIBAN_SELECT_NUMBER = 50
MAX_CHIBAN_INTERVAL = 10
MAX_WAIT_TIME = 10 # 最大待ち時間、単位：秒
CHIBAN_RETRY_WAIT_TIME = 5 # 地番・家屋番号一覧表示で再実行が必要な場合の待ち時間、単位：秒
CHIBAN_RETRY_OUT_COUNT = 5 # 地番・家屋番号一覧検索のリトライアウト数、5回連続して検索エラーなら処理中断

import pytz
tokyo_tz = pytz.timezone('Asia/Tokyo')

# 実行時間確認
#（1）土曜日及び日曜日並びに国民の祝日に関する法律（昭和23年法律第178号）に規定する休日（以下「休日」といいます。）を除いた日
#   午前8時30分から午後11時までの間（地図及び図面については午前8時30分から午後9時までの間）
#（2）土曜日及び日曜日並びに休日
#   午前8時30分から午後6時までの間（地図及び図面を除きます。） ＝ 終日実行できない
import jpholiday
import locale
locale.setlocale(locale.LC_TIME, 'ja_JP.UTF-8')


def is_RunEnable(date):
    # return True
    # 祝日は実行不可
    if jpholiday.is_holiday(date):
        print('祝日は実行不可')
        return False
    # 土は実行不可
    if date.weekday() == 5:
        print('# 土は実行不可')
        return False
    # 日は実行不可
    if date.weekday() == 6:
        print('日は実行不可')
        return False
    #  午前8時30分以前は実行不可
    if int(date.strftime('%H')) <= 8 and int(date.strftime('%M')) <= 30:
        print('午前8時30分以前は実行不可')
        return False
    #  午後9時以降は実行不可
    if int(date.strftime('%H')) >= 21:
        print('午後9時以降は実行不可')
        return False
    
    return True

from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ログイン
def login(ctrller, userId, password):
    # ログイン画面表示
    ctrller.getUrl('https://www.touki.or.jp/TeikyoUketsuke/')

    # ログイン画面の表示待ち＝「ID番号」の表示待ち
    ctrller.wait(MAX_WAIT_TIME, By.ID, 'userId')

    # 「ID番号」の入力
    ctrller.send_keys(By.ID, 'userId', userId, hide_input_value=True)  # ★自分のID番号に書き替えてください★

    # 「パスワード」の入力
    ctrller.send_keys(By.ID, 'password', password, hide_input_value=True)  # ★自分のパスワードに書き替えてください★

    # 「ログイン」ボタンをクリック		
    ctrller.click(By.XPATH, '//button/span[text()="ログイン"]')

    # 「強制ログイン」ボタン、「不動産請求」タブ、「不動産請求」リンクのいずれかの表示待ち
    ctrller.wait_any_of(MAX_WAIT_TIME,
                        By.XPATH, '//*[@id="mainBox"]//span[text()="強制ログイン"]',
                        By.XPATH, '//*[@id="uketsukeMenuFrom"]//span[text()="不動産請求"]',
                        By.XPATH, '//a/span[text()="不動産請求"]',
                        By.XPATH, '//*[@id="errorMessageArea"]/ul/li/span')
    
    # ログインエラーのとき、例外発生
    if ctrller.is_enabled(By.XPATH, '//*[@id="errorMessageArea"]/ul/li/span'):
        error_message_login =ctrller.get_text(By.XPATH, '//*[@id="errorMessageArea"]/ul/li/span')
        if "ID番号又はパスワードに誤りがあります。" in error_message_login:
            raise Exception("ID番号又はパスワードに誤りがあります。")
        else:
            raise Exception(error_message_login)
    
    # 「強制ログイン」ボタンがあれば、クリック
    if ctrller.is_enabled(By.XPATH, '//*[@id="mainBox"]//span[text()="強制ログイン"]'):
        ctrller.click(By.XPATH, '//*[@id="mainBox"]//span[text()="強制ログイン"]')

    # 「不動産請求」タブ、「不動産請求」リンクのいずれかの表示待ち
    ctrller.wait_any_of(MAX_WAIT_TIME,
                        By.XPATH, '//*[@id="uketsukeMenuFrom"]//span[text()="不動産請求"]',
                        By.XPATH, '//a/span[text()="不動産請求"]'
                        )

# ログアウト
def logout(ctrller):
    # 「ログアウト」できるときにログアウトする
    if ctrller.get_element_count(By.XPATH, '//*[@id="CHeader"]//span[contains(text(),"ログアウト")]') > 0:
        # 「ログアウト」がクリッカブルでないときは「キャンセル」をクリックする ★ToDo★
        ctrller.click(By.XPATH, '//*[@id="CHeader"]//span[contains(text(),"ログアウト")]')
        ctrller.wait(MAX_WAIT_TIME, By.XPATH, '//*[@id="mainBox"]//p[contains(text(),"ご利用ありがとうございました。")]')

# 収集条件チェック（収集条件）
def checkConditions(conditions):
    # 例）【土地/建物】土地／【都道府県番号】32／【市町村名】松江市東奥谷町／【地番・家屋番号】380～389／【seikyuJiko】全部事項
    if len(conditions) != 6:
        print(f'収集条件[{conditions}]の個数[{len(conditions)}]に過不足がある。正しくは6個')
        return False
    shozaiType = conditions[0]
    if not(shozaiType == '土地' or shozaiType == '建物'):
        print(f'収集条件の種別[{shozaiType}]が「土地」「建物」でない')
        return False
    todofukenShozai = conditions[1]
    if type(todofukenShozai) != str:
        print(f'収集条件の都道府県番号[{todofukenShozai}]が文字列でない')
        return False
    if len(todofukenShozai) != 2:
        print(f'収集条件の都道府県番号[{todofukenShozai}]が2桁でない')
        return False
    if todofukenShozai.isdigit() == False:
        print(f'収集条件の都道府県番号[{todofukenShozai}]が数値でない')
        return False
    if int(todofukenShozai) < 1 or int(todofukenShozai) > 47:
        print(f'収集条件の都道府県番号[{todofukenShozai}]が01～47でない')
        return False
    chibanKuiki = conditions[2]
    if len(chibanKuiki) == 0:
        print(f'収集条件の市町村名[{chibanKuiki}]がない')
        return False
    chiban_from = conditions[3]
    if chiban_from.isdigit() == False:
        print(f'収集条件の地番・家屋番号の開始[{chiban_from}]が数値でない')
        return False
    chiban_to = conditions[4]
    if chiban_to.isdigit() == False:
        print(f'収集条件の地番・家屋番号の終了[{chiban_to}]が数値でない')
        return False
    if int(chiban_from) > int(chiban_to):
        print(f'収集条件の地番・家屋番号の開始[{chiban_from}]と終了[{chiban_to}]の大小関係が正しくない')
        return False
    seikyuJiko_list = conditions[5]
    if type(seikyuJiko_list) != list:
        print(f'収集条件の請求種別[{seikyuJiko_list}]の指定が正しくない。配列で指定すること')
        return False
    if len(seikyuJiko_list) == 0:
        print(f'収集条件の請求種別[{seikyuJiko_list}]がない')
        return False
    for seikyuJiko in seikyuJiko_list:
        if not(seikyuJiko == '全部事項' or 
               seikyuJiko == '土地所在図/地積測量図' or 
               seikyuJiko == '建物図面/各階平面図'):
            print(f'収集条件の請求種別[{seikyuJiko}]が「全部事項」「土地所在図/地積測量図」「建物図面/各階平面図」でない')
            return False
        if shozaiType == '土地' and seikyuJiko == '建物図面/各階平面図':
            print(f'収集条件の種別[{shozaiType}]には請求種別[{seikyuJiko}]は指定できない')
            return False
        if shozaiType == '建物' and seikyuJiko == '土地所在図/地積測量図':
            print(f'収集条件の種別[{shozaiType}]には請求種別[{seikyuJiko}]は指定できない')
            return False
    return True


# 検索画面表示
def displaySearchScreen(ctrller):
    # マイページのリンクがあればクリック
    if ctrller.is_enabled(By.XPATH, '//a/span[text()="マイページ"]'):
        ctrller.click(By.XPATH, '//a/span[text()="マイページ"]')
        ctrller.wait(MAX_WAIT_TIME, By.XPATH, '//a/span[text()="不動産請求"]')

    # 不動産請求のリンク/タブのクリックから画面の表示まで
    # 「不動産請求」タブがあればクリック
    if ctrller.is_enabled(By.XPATH, '//a/span[text()="不動産請求"]'):
        ctrller.click(By.XPATH, '//a/span[text()="不動産請求"]')

    # 「不動産請求」リンクがあればクリック
    if ctrller.is_enabled(By.XPATH, '//*[@id="uketsukeMenuFrom"]//span[text()="不動産請求"]'):
        ctrller.click(By.XPATH, '//*[@id="uketsukeMenuFrom"]//span[text()="不動産請求"]')    

    # 「不動産請求」タブの中身の表示待ち＝「所在－都道府県」の表示待ち
    ctrller.wait(MAX_WAIT_TIME, By.ID, 'fuTodofukenShozai')


# 検索条件入力(収集条件)
def enterCondition(ctrller, conditions):
    # 例）【土地/建物】土地／【都道府県番号】32／【市町村名】松江市東奥谷町／【地番・家屋番号】380～389／【請求種別】全部事項
    shozaiType = conditions[0]
    todofukenShozai = conditions[1]
    chibanKuiki = conditions[2]
    seikyuJiko_list = conditions[5]

    # 種別
    if shozaiType == '土地':
        ctrller.click(By.ID, 'fuShozaiTypeTOCHI')
    if shozaiType == '建物':
        ctrller.click(By.ID, 'fuShozaiTypeTATEMONO')

    # 「都道府県」の選択－「01:北海道」～「47:沖縄県」を選択
    ctrller.select(By.ID, 'fuTodofukenShozai', todofukenShozai)

    # 「直接入力」が選択状態でなければクリック
    if ctrller.driver.find_element(By.ID, "fuShozaiChokusetuNyuryoku").is_selected() == False:
        ctrller.click(By.ID, 'fuShozaiChokusetuNyuryoku')

    # 「直接入力」を入力
    ctrller.send_keys(By.ID, 'fuChibanKuiki', chibanKuiki)

    # 「請求種別」の選択は一旦クリア
    for seikyuJiko_ID in ['fuAll', 'fuShoyusya', 'fuChizu', 'fuShozai','fuChieki', 'fuZumen']:
        if ctrller.driver.find_element(By.ID, seikyuJiko_ID).is_selected() == True:
            ctrller.click(By.ID, seikyuJiko_ID)

    # 「請求種別」の選択
    for seikyuJiko in seikyuJiko_list:
        if seikyuJiko == "全部事項":
            ctrller.click(By.ID, 'fuAll')
        if seikyuJiko == "土地所在図/地積測量図":
            ctrller.click(By.ID, 'fuShozai')
        if seikyuJiko == "建物図面/各階平面図":
            ctrller.click(By.ID, 'fuZumen')

    return


# 地番選択(取得開始位置)
#   地番選択50件ごとに以下の処理呼び出し ※取得位置の初期値は1
#   ※戻り値
#      0 ： すべての地番を選択できた場合
#     -1 ： 選択すべき地番がなかった場合
#     取得開始位置=start_select_number + 50 ： 選択すべき地番が残っている場合の次の取得開始位置
def selectChiban(ctrller, xlsCtr, start_select_number, chiban_from, chiban_to):

    next_start_select_number = -1  # 地番選択なし（初期値）

    # 「地番・家屋番号」をクリック
    ctrller.wait(MAX_WAIT_TIME, By.ID, 'fuChibanKaokuIchiran')
    ctrller.click(By.ID, 'fuChibanKaokuIchiran')
    
    # 「地番・家屋番号選択」画面の表示待ち＝「所在－検索範囲」の表示待ち
    ctrller.wait(MAX_WAIT_TIME, By.ID, 'cbnDlgSearchChibanStart')
    
    # 「所在－検索範囲」の入力
    ctrller.send_keys(By.ID, 'cbnDlgSearchChibanStart', chiban_from)
    ctrller.send_keys(By.ID, 'cbnDlgSearchChibanEnd', chiban_to)
    # ToDo：ここでフォーカスアウトするとよいのかも？ ⇒ 効果なし
    ctrller.focusToElement(By.ID, 'cbnDlgChibanSearch')

    # 地番・家屋番号の再検索回数
    chiban_retry_count = 0

    while True:
        # 「検索」ボタンをクリック
        time.sleep(CHIBAN_RETRY_WAIT_TIME * (1 + 0.5 * chiban_retry_count) )  # 頻繁に再実行になるため、ここに移動、3回でもエラーになるので、間隔を徐々に空けてみる
        ctrller.click(By.ID, 'cbnDlgChibanSearch')
        
        # 地番・家屋番号一覧の表示待ち
        ctrller.wait_any_of(MAX_WAIT_TIME,
                            By.ID, 'cbnDlgChibanChk_1',
                            By.XPATH, '//*[@id="cbnDlgErrMsgArea"]/ul/li'
        )

        # 検索結果なしは //*[@id="cbnDlgErrMsgArea"]/ul/li が存在する
        if ctrller.is_enabled(By.XPATH, '//*[@id="cbnDlgErrMsgArea"]/ul/li'):
            errorText = ctrller.get_text(By.XPATH, '//*[@id="cbnDlgErrMsgArea"]/ul/li')
            # 再検索が必要な場合、
            if 'しばらく時間を空けてから、再度実行してください。' in errorText:
                # 頻繁に再実行になるため、「検索」ボタンのクリック前に移動
                # time.sleep(CHIBAN_RETRY_WAIT_TIME)
                # {HIBAN_RETRY_OUT_COUNT}回再実行してもエラーなら、処理中断
                if chiban_retry_count == (CHIBAN_RETRY_OUT_COUNT - 1):
                    # エラーメッセージ出力
                    errorMessage = f'「{chiban_from}」～「{chiban_to}」の条件にて地番・家屋番号の検索ができません。\n'
                    errorMessage += f'{CHIBAN_RETRY_OUT_COUNT}回連続して検索エラーとなりました。\n処理を中断します。'
                    print(errorMessage)
                    # 処理中断をExcelに出力
                    # xlsCtr.save_with_exception()
                    g_process_info['status'] = False
                    g_process_info['message'] = errorMessage
                    if g_isDisplayMessage:
                        Message.MessageForefrontShowinfo(errorMessage)
                    # 「キャンセル」ボタンをクリック
                    ctrller.click(By.ID, 'cbnDlgBtnCancel')
                    return next_start_select_number
                
                # 地番・家屋番号の再検索回数のカウントアップ
                chiban_retry_count += 1
                
                # 検索再実行
                continue

            # 検索結果なしの場合、正常扱い、処理継続
            elif ('指定した範囲内に登記情報がありません。指定範囲を変更して請求してください。' in errorText):
                # 「キャンセル」ボタンをクリック
                ctrller.click(By.ID, 'cbnDlgBtnCancel')
                return next_start_select_number
            
            elif ('請求できない所在です' in errorText): # ★ToDo★ OR条件で記載すること
                # 「キャンセル」ボタンをクリック
                ctrller.click(By.ID, 'cbnDlgBtnCancel')
                return next_start_select_number
            
            else:
                # エラーメッセージ出力
                errorMessage = '「{chiban_from}」～「{chiban_to}」の条件にて地番・家屋番号の選択ができません。\n'
                errorMessage = '処理を中断します。'
                print(errorMessage)
                # 処理中断をExcelに出力
                # xlsCtr.save_with_exception()
                g_process_info['status'] = False
                g_process_info['message'] = errorMessage
                if g_isDisplayMessage:
                    Message.MessageForefrontShowinfo(errorMessage)
                # 「キャンセル」ボタンをクリック
                ctrller.click(By.ID, 'cbnDlgBtnCancel')
                return next_start_select_number
        
        else:
            break

    # チェックボックスのクリック ※先頭から最大５０個クリック
    id_num = 1
    page = 1

    # 地番・家屋番号の選択を一旦クリア
    if ctrller.is_enabled(By.ID, 'cbnDlgBtnCheckAllReset'):           
        ctrller.click(By.ID, 'cbnDlgBtnCheckAllReset')

    # 地番・家屋番号の開始位置から最大地番選択数を選択する
    while True:
        if id_num == 101: # 1ページの表示件数はMAX100件。「次」ボタンをクリックする必要あり
            # 次が有効でないなら選択なしで終了
            if ctrller.is_enabled(By.ID, 'cbnDlgBtnPageNext') == False:
                next_start_select_number = -1
                break
            
            # 「次」ボタンをクリックし、地番・家屋番号一覧の次ページを表示
            ctrller.click(By.ID, 'cbnDlgBtnPageNext')
            
            # 一覧表示待ち ★ToDo★ 再検索が必要なるケースあり
            ctrller.wait(MAX_WAIT_TIME, By.ID, 'cbnDlgChibanChk_1')

            id_num = 1
            page += 1

        id_check = f'cbnDlgChibanChk_{id_num}'
        read_num = id_num + (page - 1) * 100
        
        # 今回、取得すべき地番・家屋番号の場合（初回は1から50番、2回目は51番から100番、3回目は101番から150番…）
        if read_num >= start_select_number and read_num <= (start_select_number + MAX_CHIBAN_SELECT_NUMBER - 1):
            if ctrller.is_enabled(By.ID, id_check):
                ctrller.click(By.ID, id_check)
                next_start_select_number = read_num
            else:
                next_start_select_number = 0 # すべての地番を選択した場合
                break

        # 次回、取得すべき地番・家屋番号が存在する場合（初回は51番、2回目は101番、3回目は151番…）
        if read_num == (start_select_number + MAX_CHIBAN_SELECT_NUMBER):
            if ctrller.is_enabled(By.ID, id_check):
                next_start_select_number = read_num
                break
            else:
                next_start_select_number = 0 # すべての地番を選択した場合
                break
        
        # 地番・家屋番号が存在しない場合
        if ctrller.is_enabled(By.ID, id_check) == False:
            if read_num > (start_select_number + MAX_CHIBAN_SELECT_NUMBER):
                next_start_select_number = -1 # 選択できなかった場合
            break
        
        id_num += 1

    # 「確定」ボタンをクリックし、条件入力画面へ地番・家屋番号を反映する
    ctrller.click(By.ID, 'cbnDlgBtnOk')
    
    # 次回取得開始位置を返却
    return next_start_select_number


# 検索結果画面（「不動産請求一覧」画面）表示
def dispalySearchResult(ctrller, xlsCtr):
    # 検索結果の表示待ち
    ctrller.wait(MAX_WAIT_TIME, By.ID, f'seikyuJiko_1')

    # 検索結果の明細ごとに以下の処理実行
    i = 1
    while (ctrller.get_element_count(By.ID, f'seikyuJiko_{i}') == 1):
        # 請求種別、種別、所在及び地番・家屋番号 / 不動産番号 を取得
        seikyuJiko = ctrller.get_text(By.ID, f'seikyuJiko_{i}')
        shozaiType = ctrller.get_text(By.ID, f'shozaiType_{i}')
        shozai     = ctrller.get_text(By.ID, f'shozai_{i}')

        # 図面一覧ボタンあり場合
        if ctrller.get_element_count(By.ID, f'jiken_{i}') == 1:
            # 「図面一覧」ボタンをクリック
            ctrller.click(By.ID, f'jiken_{i}')

            # 「図面事件一覧」画面の表示待ち＝一覧表orエラーメッセージの表示待ち
            ctrller.wait_any_of(MAX_WAIT_TIME,
                                By.XPATH, '//table[@id="jikenListTbl"]//td[@class="col_w2"]',
                                By.XPATH, '//*[@id="jkDlgErrMsgArea"]/ul/li[1]')
                
            # データの有無を判断
            cnt_zumen = ctrller.get_element_count(By.XPATH, '//table[@id="jikenListTbl"]//td[@class="col_w2"]')
            data_zumen_list = []
            for j in range(0, cnt_zumen):
                data_zumen = {
                    '登録年月日' : ctrller.get_text(By.XPATH, '//table[@id="jikenListTbl"]//td[@class="col_w2"]', j),
                    '事件ID' : ctrller.get_text(By.XPATH, '//table[@id="jikenListTbl"]//td[@class="col_w3"]', j),
                    '所在及び地番・家屋番号（事件前物件）' : ctrller.get_text(By.XPATH, '//table[@id="jikenListTbl"]//td[@class="col_w4"]', j),
                    '所在及び地番・家屋番号（事件後物件）' : ctrller.get_text(By.XPATH, '//table[@id="jikenListTbl"]//td[@class="col_w5"]', j),
                }
                data_zumen_list.append(data_zumen)

            if cnt_zumen > 0:
                for data_zumen in data_zumen_list:
                    # 不動産請求明細＋図面情報データ出力
                    xlsCtr.writeZumen(seikyuJiko, shozaiType, shozai,
                                      data_zumen['登録年月日'], 
                                      data_zumen['事件ID'], 
                                      data_zumen['所在及び地番・家屋番号（事件前物件）'], 
                                      data_zumen['所在及び地番・家屋番号（事件後物件）'] )
            else:
                # 不動産請求明細＋図面なしの理由のデータ出力
                reasonZumenNashi = ctrller.get_text(By.XPATH, '//*[@id="jkDlgErrMsgArea"]/ul/li[1]')
                xlsCtr.writeZemenNasi(seikyuJiko, shozaiType, shozai, reasonZumenNashi)

            # 「キャンセル」ボタンをクリックし、「不動産請求一覧」画面に戻る
            ctrller.click(By.ID, 'jkDlgBtnCancel')
           
        # 図面一覧ボタンなし場合
        else:
            # 不動産請求明細データ出力
            xlsCtr.writeFudousan(seikyuJiko, shozaiType, shozai)

        # 次の行
        i += 1


import traceback
g_isDisplayMessage = True # メッセージ表示=既定値：True
# 収集処理状態
g_process_info = {
            'status'  : True, # True：正常終了／False：異常終了
            'message' : None,
            'html' : None,
            'traceback' : None
        }
        
# テスト用スタブ：データ収集（収集条件）異常終了
def collectData_stab_abnormal(conditions, user_id, password, isDisplayMessage=True):
    errorMessage  = f'収集条件：{xlsContorller.editCollectionCondition(conditions)}の収集処理にてエラーが発生しました。\n'
    errorMessage += f'この処理を中断します。エラー対処後、再実行してください。'
    g_process_info = {
        'status'  : False, # True：正常終了／False：異常終了
        'message' : errorMessage,
        'html' : '<html>～</html>',
        'traceback' : 'traceback'
    }
    # print(f'g_process_info={g_process_info}')
    time.sleep(3)
    return r'.\output\output_20241115_163708.xlsx', g_process_info

# テスト用スタブ：データ収集（収集条件）正常終了
def collectData_stab(conditions, user_id, password, isDisplayMessage=True):
    # print(f'g_process_info={g_process_info}')
    time.sleep(3)
    return r'.\output\output_20241115_163708.xlsx', g_process_info


# データ収集（収集条件）
def collectData(conditions, user_id, password, isDisplayMessage=True, isCloud=False):
    ctrller = None
    xlsCtr = None
    g_isDisplayMessage = isDisplayMessage
    g_process_info = {
            'status'  : True, # True：正常終了／False：異常終了
            'message' : None,
            'html' : None,
            'traceback' : None
        }
    try:
        # データ出力コントローラー生成
        if isCloud == False:
            xlsCtr = xlsContorller.xlsContorller()
        else:
            xlsCtr = xlsContorller.gsContorller()

        # 収集条件のデータ出力
        xlsCtr.writeCondition(user_id, conditions)
        
        # ブラウザ起動
        ctrller = selenimuContorller.selenimuContorller(isCloud=True)

        # ログイン(ログイン、パスワード)
        login(ctrller, user_id, password)
        
        # 地番・家屋番号開始、終了の10件ごとに以下の処理呼び出し
        # 地番・家屋番号の間隔を10：最大地番間隔で割って切り捨て
        chiban_from = int(conditions[3])
        chiban_to   = int(conditions[4])
        d = (chiban_to - chiban_from) // MAX_CHIBAN_INTERVAL

        # 地番・家屋番号を10個づつ分ける
        for i in range(d + 1):
            f = chiban_from + i * MAX_CHIBAN_INTERVAL
            t = f + MAX_CHIBAN_INTERVAL - 1
            if t > chiban_to:
                t = chiban_to
            # print(f'{i+1}回目：From:{f} To:{t}')

            # 検索条件入力画面表示
            displaySearchScreen(ctrller)

            # 検索条件入力
            conditions_division = conditions
            conditions_division[3] = str(f)
            conditions_division[4] = str(t)
            enterCondition(ctrller, conditions_division)

            start_select_number = 1
            while True:
                # 地番・家屋番号をMAX50件ごと選択する
                next_start_select_number = selectChiban(ctrller, xlsCtr, start_select_number, conditions_division[3], conditions_division[4])
                start_select_number = next_start_select_number

                # 地番・家屋番号の選択なし
                if start_select_number < 0:
                    break
                
                # 「確定」ボタンをクリックし、「不動産請求一覧」画面を表示
                ctrller.click(By.XPATH, '//*[@id="tabsFudosan"]//span[contains(text(),"確定")]')
                
                # 検索結果画面（「不動産請求一覧」画面）表示
                dispalySearchResult(ctrller, xlsCtr)

                # データセーブ
                # xlsCtr.save()

                # 「戻る」ボタンをクリック
                ctrller.click(By.XPATH, '//span[contains(text(),"戻る")]')

                # 次回取得データがない場合
                if start_select_number == 0:
                    break

        # データセーブ
        xlsCtr.save()

        # ログアウト
        logout(ctrller)
        

    except Exception as e:
        # エラー発生をExcelに出力
        if xlsCtr != None:
            xlsCtr.save_with_exception()

        # ログインエラーのとき
        if 'ID番号又はパスワードに誤りがあります。' in f'{e}':
            # エラーメッセージ
            errorMessage  = f'ID番号又はパスワードに誤りがあります。'
            print(errorMessage)

        # それ以外のとき
        else:
            # エラーメッセージ
            errorMessage  = f'収集条件：{xlsContorller.editCollectionCondition(conditions)}の収集処理にてエラーが発生しました。\n'
            errorMessage += f'この処理を中断します。エラー対処後、再実行してください。'
            print(errorMessage)

            # 例外発生時のトレースバック、HTMLソース、画面スナップショットの出力           
            print(f"≫≫≫≫≫ トレースバック情報 ここから ≫≫≫≫≫\n")
            g_process_info['traceback'] = traceback.format_exc()
            print(f"{traceback.format_exc()}\n")
            print(f"≪≪≪≪≪ トレースバック情報 ここまで ≪≪≪≪≪")

            if ctrller != None:
                print(f"≫≫≫≫≫ HTMLソース ここから ≫≫≫≫≫\n")
                g_process_info['html'] = ctrller.driver.page_source
                print(f"{ctrller.driver.page_source}\n")
                print(f"≪≪≪≪≪ HTMLソース ここまで ≪≪≪≪≪")

                # スナップショット,
                # # Cloud対応の場合、スナップショットは取らない
                # if isCloud == False:
                ctrller.driver.set_window_size(1048, 1048)
                ctrller.driver.get_screenshot_as_file(f".\\output\\snapshot_{datetime.now(tokyo_tz).strftime("%Y%m%d_%H%M%S")}.png")
                
        # エラーメッセージ表示
        g_process_info['status'] = False
        g_process_info['message'] = errorMessage
        if g_isDisplayMessage:
            Message.MessageForefrontShowinfo(errorMessage)

    # ブラウザ閉じる
    ctrller.close()
    
    # 収集結果ファイル名を返却
    if isDisplayMessage:
        return xlsCtr.output_file_path
    else:
        return xlsCtr.output_file_path, g_process_info
