import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import os
import webbrowser
import pandas as pd
from flask import Flask, render_template, request


app = Flask(__name__)

@app.route('/')
def spreadsheet_view():
    # use creds to create a client to interact with the Google Drive API 認証情報
    scope = ['https://spreadsheets.google.com/feeds'] #google spreadsheet のAPI
    creds = ServiceAccountCredentials.from_json_keyfile_name('spread-sheet-386510-6b8a91b72e28.json', scope) #秘密鍵json file名
    client = gspread.authorize(creds) #authenticate with Google

    # spreadsheatを指定してsheet1を開く
    spreadsheet_id = '1Lo2XlNUG97LDFjjugDh1CRPDEWzp8x-AlYaECOIRRXE'  # Please set your Google Sheets ID
    sheet = client.open_by_key(spreadsheet_id).sheet1

    # 今日の日付を取得(YY/MM/DD形式)
    today = datetime.date.today().strftime("%Y/%m/%d")

    #全ての行を取得
    rows = sheet.get_all_values()


    # J列とK列が空白である行を保存するリスト
    empty_rows = []
    # 各行をループし、J列とK列が空白である行のA列の内容とその行番号を保存
    for i, row in enumerate(rows, start=1):  #rowはある行を1D配列として格納したもの.
        if not row[9] or not row[10]:  # 列は0から始まるので、J列とK列はそれぞれ9と10
            empty_rows.append((i, row))


    # 空白の行を一覧表示
    empty_rowlist = []
    index_list = []
    for row_num, row in empty_rows:
        dummy = row
        empty_rowlist.append(dummy)
        index_list.append(row_num)

    df = pd.DataFrame(empty_rowlist,index = index_list,columns=['通称', '型番', 'Web', '予算管理者', '使用者', '単価', '数量', '合計', '発注日', '納品日', '伝票処理日', '定価'])

    #htmlテーブルに変換
    data = []
    for row_num, row in empty_rows:
        data.append([row_num] + row)  # 行番号を行の先頭に追加
    return render_template('home.html', data=data)



@app.route('/update', methods=['POST'])
def update_spreadsheet():
    # ユーザーから送信されたデータを取得
    row_num = request.form.get('row_num')
    company = request.form.get('company')
    date_choice = request.form.getlist('date_choice')

    # 取得したデータを表示（確認用）
    print(f"row_num: {row_num}")
    print(f"company: {company}")
    print(f"date_choice: {date_choice}")

    # 今日の日付を取得
    today = datetime.date.today().strftime("%Y%m%d")

    # Google Spreadsheetsにアクセス
    scope = ['https://spreadsheets.google.com/feeds']
    creds = ServiceAccountCredentials.from_json_keyfile_name('spread-sheet-386510-6b8a91b72e28.json', scope)
    client = gspread.authorize(creds)
    spreadsheet_id = '1Lo2XlNUG97LDFjjugDh1CRPDEWzp8x-AlYaECOIRRXE'
    sheet = client.open_by_key(spreadsheet_id).sheet1

    # 選択された行を更新
    for choice in date_choice:
        if choice == 'delivery_date':
            sheet.update_cell(int(row_num), 10, today)  # 10 is the column number for "納品日"
        elif choice == 'invoice_processing_date':
            sheet.update_cell(int(row_num), 11, today)  # 11 is the column number for "伝票処理日"

    # ディレクトリを作成
    row_values = sheet.row_values(int(row_num))
    nickname = row_values[0]  # assuming that "通称" is in the first column
    directory_name = f"C:/Users/mikku/MyFiles/test/{today}_{company}_{nickname}"#f関数はかっこ内の変数をそのまま埋め込める．
    os.makedirs(directory_name, exist_ok=True)

    return render_template('update_successful.html')


