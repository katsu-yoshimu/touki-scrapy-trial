import streamlit as st
import json
import xlsContorller
import toukiController
from datetime import datetime
import time
from preflist import PREF_CODE

CONFIG_FILE = 'config.json'
IS_CLOUD = True

# 設定ファイル読込
def readConfig():
    config = None

    # 初回呼び出しのとき（セッションが存在しないとき）、設定ファイルから初期表示データを取得
    if "read_config" not in st.session_state:
        # print("設定ファイルから初期表示データを取得")
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as file:
                config = json.load(file)
        except Exception as e:
            errorMessage=f'設定ファイルの読み込みでエラーが発生しました。\nエラー内容[{e}]'
            print(errorMessage)
            # 初期値：ファイルがないとき
            config = {
                "user_id"  : "XXXXXX",
                "password" : "XXXXXX",
                "conditions_list" : [
                    ["土地", "32", "松江市東奥谷町", "380", "389", ["全部事項"]]
                ]
            }
        # 設定ファイル読込済をセッションに記録
        st.session_state['read_config'] = True
    
    # ２回目以降の呼び出しのとき（セッションが存在するとき）、セッションから初期表示データを取得
    else:
        # print("セッションから初期表示データを取得")
        config = {
            "user_id"  : st.session_state.get('user_id'),
            "password" : st.session_state.get('password'),
            "conditions_list" : [
                [
                    st.session_state.get('shozaiType'),
                    PREF_CODE[st.session_state.get('todofukenShozai')],
                    st.session_state.get('chibanKuiki'),
                    str(st.session_state.get('chiban_from')),
                    str(st.session_state.get('chiban_to')),
                    st.session_state.get('seikyuJiko')
                ]
            ]
        }

    return config

# 入力フォーム表示
def dispalyForm(config):
    st.text('「不動産請求情報」収集条件 入力')

    # submitボタンでセッションにデータを保持するため入力フォームに「key」を指定
    # submitボタンを押すまでリロードしないように「enter_to_submit=False」を指定
    with st.form(key="my_form", enter_to_submit=False):
        # ID番号、パスワード
        col1, col2 = st.columns((1, 1))
        with col1:
            user_id  = st.text_input('ID番号',   config['user_id'], key='user_id')
        with col2:
            password = st.text_input('パスワード', config['password'], key='password')

        # 種別
        if '土地' == config['conditions_list'][0][0]:
            shozaiType_index = 0
        else:
            shozaiType_index = 1
        shozaiType = st.radio('種別', ['土地', '建物'], key='shozaiType', horizontal=True, index=shozaiType_index)  # index=1のとき'建物'が選択された状態となる

        # 都道府県、市町村名
        col1, col2 = st.columns((1, 4))
        with col1:
            todofukenShozai = st.selectbox('都道府県',  PREF_CODE.keys(), key='todofukenShozai', index=(int(config['conditions_list'][0][1])-1))
        with col2:
            chibanKuiki = st.text_input('市町村名', config['conditions_list'][0][2], key='chibanKuiki')

        # 地番・家屋番号
        col1, col2 = st.columns((1, 1))
        with col1:
            chiban_from = st.number_input('地番・家屋番号 開始',  key='chiban_from', value=int(config['conditions_list'][0][3]))
        with col2:
            chiban_to   = st.number_input('地番・家屋番号 終了',  key='chiban_to',   value=int(config['conditions_list'][0][4]))

        # 請求種別
        options = ["全部事項", "土地所在図/地積測量図", "建物図面/各階平面図"]
        seikyuJiko = st.pills("請求種別", options, default=config['conditions_list'][0][5], selection_mode="multi", key='seikyuJiko')

        # 実行ボタン
        submit_button = st.form_submit_button("収集実行")

# 入力チェック
def checkForm(config):
    isCheck = True

    user_id     = config['user_id'] 
    password    = config['password']
    shozaiType  = config['conditions_list'][0][0]
    chibanKuiki = config['conditions_list'][0][2]
    chiban_from = int(config['conditions_list'][0][3])
    chiban_to   = int(config['conditions_list'][0][4])
    seikyuJiko  = config['conditions_list'][0][5]

    # ID番号チェック
    if user_id == '':
        st.warning("ID番号が未入力です。") 
        isCheck = False

    # パスワードチェック
    if password == '':
        st.warning("パスワードが未入力です。") 
        isCheck = False

    # 市町村名チェック
    # 市町村名で検索エラーになるケースは収集結果Zero件となるためチェックなし
    if chibanKuiki == '':
        st.warning("市町村名が未入力です。") 
        isCheck = False

    # 地番・家屋番号チェック
    if chiban_from <= 0 or chiban_to <= 0:
        st.warning("地番・家屋番号は「1」以上を入力してください。") 
        isCheck = False    
    if chiban_from > chiban_to:
        st.warning("地番・家屋番号の開始、終了の大小関係に誤りがあります。") 
        isCheck = False
    
    # 請求種別チェック
    if shozaiType == "土地" and '建物図面/各階平面図' in seikyuJiko:
        st.warning("種別が「土地」のとき、請求種別に「建物図面/各階平面図」は選択できません。") 
        isCheck = False
    if shozaiType == "建物" and '土地所在図/地積測量図' in seikyuJiko:
        st.warning("種別が「建物」のとき、請求種別に「土地所在図/地積測量図」は選択できません。") 
        isCheck = False
    if len(seikyuJiko) == 0:
        st.warning("請求種別が未選択です。いずれかを選択してください。") 
        isCheck = False

    # 実行時間確認
    if toukiController.is_RunEnable(datetime.now()) == False:
        errorMessage = f'現在は{datetime.now().strftime('%Y/%m/%d %H:%M:%S')}です。実行時間外です。\n平日8時30分～21時0分に実行してください。'
        st.warning(errorMessage)
        isCheck = False

    # いずれかエラーがある場合は表示終了    
    if isCheck == False:
        st.stop()

# 収集処理の実行確認ダイアログ表示
@st.dialog("収集実行確認")
def approve_run(config):

    user_id = config['user_id'] 
    password = config['password']
    conditions = config['conditions_list'][0]

    # 収集開始メッセージ
    st.text("不動産請求情報収集処理を実行しますか？")
    st.text(f'実行ユーザはID番号【{user_id}】、パスワード【{password}】です。\n')
    st.text(f'収集条件は以下の通りです。')
    st.text(xlsContorller.editCollectionCondition(conditions))

    # 「OK」「Cancel」ボタン表示
    col1, col2 = st.columns((1, 1))
    with col1:
        OK      = st.button("OK")
    with col2:
        Cancel  = st.button("Cancel")
    
    # 「OK」クリックのとき、セッションに記録
    if OK:
        st.session_state["ok_approve_run"] = True
        st.rerun()
    if Cancel:
        st.rerun()


# 収集処理実行
def run(config):

    user_id = config['user_id'] 
    password = config['password']
    conditions = config['conditions_list'][0]

    # 設定ファイルに書き戻す
    with open(CONFIG_FILE, 'w', encoding='utf-8') as file:
       json.dump(config, file, indent=4, ensure_ascii=False)

    # 収集処理実行
    # 処理終了時のメッセージ表示のため、出力ファイル名を追記
    output_file_path, process_info = toukiController.collectData(conditions, user_id, password, isDisplayMessage=False, isCloud=IS_CLOUD)

    # 収集終了メッセージ
    # 正常終了時
    if process_info['status']:
        endMessage = '不動産請求情報収集処理が終了しました。\n\n'
        endMessage += f'収集結果は【 {output_file_path} 】に出力されています。ご確認ください。'
        st.info(endMessage)

    # 異常終了時
    else:
        endMessage = '不動産請求情報収集の処理中に異常が発生しました。\n\n'
        endMessage += f'収集結果は【 {output_file_path} 】に出力されています。ご確認ください。'
        st.warning(endMessage)
        
        # エラーメッセージ、トレースバック情報、HTMLソースを表示
        if process_info['message']:
            st.warning(process_info['message'])

        if process_info['traceback']:
            st.code(process_info['traceback'])

        if process_info['html']:
            st.code(process_info['html'])


### ここから画面表示処理 ###
# 設定ファイル読込
config = readConfig()

# 入力フォーム表示
dispalyForm(config)

# 入力値チェック
checkForm(config)

# 実行確認ダイアログ
if "ok_approve_run" not in st.session_state:
    if st.session_state['FormSubmitter:my_form-収集実行']:
        approve_run(config)

# 実行確認ダイアログで「OK」をクリックしたとき
if st.session_state.get("ok_approve_run", False):
    # 実行確認ダイアログで「OK」をクリックのセッションをクリア
    del st.session_state['ok_approve_run']
    # 収集実行
    run(config)
