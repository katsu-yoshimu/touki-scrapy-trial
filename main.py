import Message
import xlsContorller
import toukiController
from datetime import datetime

output_file_list = []

import json

def main():
    # 設定ファイル読込
    config = None
    try:
        with open('config.json', 'r', encoding='utf-8') as file:
            config = json.load(file)
    except Exception as e:
        errorMessage=f'設定ファイルの読み込みでエラーが発生しました。\nエラー内容[{e}]'
        print(errorMessage)
        Message.MessageForefrontShowinfo(errorMessage)
        return
    
    user_id = config['user_id']
    password = config['password']
    conditions_list = config['conditions_list']

    # 実行時間確認
    if toukiController.is_RunEnable(datetime.now()) == False:
        errorMessage = f'現在は{datetime.now().strftime('%Y/%m/%d %H:%M:%S')}です。実行時間外です。\n平日8時30分～21時0分に実行してください。'
        print(errorMessage)
        Message.MessageForefrontShowinfo(errorMessage)
        return
    
    # 収集条件チェック(収集条件) 
    for conditions in conditions_list:
        if toukiController.checkConditions(conditions) == False:
            # エラーメッセージを表示し、処理終了
            errorMessage = '収集条件に誤りがあります。\nログを参照し、訂正後、再実行してください。'
            print(errorMessage)
            Message.MessageForefrontShowinfo(errorMessage)
            return

    # 収集開始メッセージ
    startMessage = f'不動産請求情報収集を実行しますか？\n実行ユーザはID番号【{user_id}】、パスワード【{password}】です。\n収集条件は以下の通りです。'
    for i, conditions in enumerate(conditions_list):
        startMessage += f'\n\n{i+1}：{xlsContorller.editCollectionCondition(conditions)}'
    print(startMessage)
    if Message.MessageForefront(startMessage) == False:
        return
    
    # データ収集
    for conditions in conditions_list:
        output_file_path = toukiController.collectData(conditions, user_id, password)
        # 処理終了時のメッセージ表示のため、出力ファイル名を追記
        output_file_list.append(output_file_path)
    
    # 収集終了メッセージ
    endMessage = '不動産請求情報収集が処理終了しました。\n収集結果は以下に出力されています。ご確認ください。'
    for i, output_file in enumerate(output_file_list):
        endMessage += f'\n{i+1}：【{output_file}】'
    print(endMessage)
    Message.MessageForefrontShowinfo(endMessage)
            
main()