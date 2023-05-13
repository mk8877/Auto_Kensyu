import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import os
import webbrowser
import pandas as pd
from flask import Flask, render_template

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
    table = df.to_html()

    #レンダリング
    return render_template('home.html', table=table)

