import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build 
from datetime import datetime, timedelta
from dotenv import load_dotenv
import os

# .envファイルから環境変数を読み込む
load_dotenv()

SERVICE_ACCOUNT_FILE = './credentials/XXX' # Credentialファイル名に修正
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
DIRECTORY_ID = os.getenv('DIRECTORY_ID')

def reading_spreadsheet(spreadsheet_key,file_name):
    spreadsheet = client.open_by_key(spreadsheet_key)
    sheet = spreadsheet.worksheet('日付') # シート名に合わせて修正
    data = sheet.get_all_values()
    # １週間前〜前日のデータ取得
    today = datetime.now().date()
    start_date = today - timedelta(days=7)
    end_date = today - timedelta(days=1)

    targetData = []
    # アカウント登録者数のindexを取得
    account_register_account_count_index = data[0].index('アカウント登録者数')
    # 初期登録者数のIndexを取得
    initial_register_account_count_index = data[0].index('初期登録者数')
    # 投票数のIndexを取得
    vote_count_index = data[0].index('投票数')
    for row in data:
        try:
            date_str = row[0]
            date_obj = datetime.strptime(date_str, '%Y/%m/%d').date()
            if start_date <= date_obj <= end_date:
                targetData.append(row)
        except ValueError:
            print('A列に日付以外の値が入っているためスキップ' + str(row))
            continue

    for row in targetData:
        if not row[account_register_account_count_index] or not row[initial_register_account_count_index] or not row[vote_count_index]:
            print(row[0]+'のデータが正しく集計されていません。　'+str(row))
            global global_hasError
            global_hasError = True
    print(file_name+'->チェック終了')

global_hasError = False
credential = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE,scopes=SCOPES)
client = gspread.authorize(credential)

drive_service = build('drive','v3',credentials=credential)
files = drive_service.files().list().execute()
spreadsheet_files = []
for file in files.get('files', []):
    # スプレッドシート以外は、除外
    if not file.get('mimeType') == 'application/vnd.google-apps.spreadsheet':
        continue
    # 利用停止物件も除外
    if '利用停止' in file.get('name'):
        continue
    spreadsheet_files.append(file)

# 取得したスプレッドシートを順番に読み込む
for file in spreadsheet_files:
    file_id = file.get('id')
    file_name = file.get('name')
    print(f"Reading spreadsheet: {file_name}")
    reading_spreadsheet(file_id,file_name)

if global_hasError:
    print("結果確認完了：集計結果NG")
else :
    print("結果確認完了：集計結果OK")