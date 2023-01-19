
import json
import os.path
# Read Spreadsheet From Google Drive

from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

creds = None

# 測試用spreadsheet id：16aWDkDS6hh0NZE6R2x-ihTlmjsnW6dzh0QLasAuZd3Q

print('You have to put your google service secret in secret/google-api-secret.json')

# Read creds json file
if os.path.exists('secret/google-api-secret.json'):
    service_account_info = json.load(open('secret/google-api-secret.json'))
    # Authentication
    creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)

# TODO: handle not valid creds

try:
    sheet_id = input('Enter google spreadsheet id\n')
    # build api base url
    service = build('sheets', 'v4', credentials=creds)

    # Call the sheets API -> 讀spread sheet
    sheet_values = service.spreadsheets().values()
    # prevent 403 you have to copy your service email in json file to google sheet share permission
    result = sheet_values.get(spreadsheetId=sheet_id, range="Sheet1").execute()
    sheets = service.spreadsheets().get(spreadsheetId=sheet_id, ranges="Sheet1").execute()
    sheet1_id = sheets['sheets'][0]['properties']['sheetId']
    rows = result.get('values', [])

    if not rows:
        print('no data found')

    # copy col title format
    batch_update_request_body = {
        'requests': [
            {
                'copyPaste': {
                    'source': {
                        'sheetId': sheet1_id,
                        'endRowIndex': 1,
                        'endColumnIndex': 1,
                    },
                    'destination': {
                        'sheetId': sheet1_id,
                        'endRowIndex': 1,
                        'endColumnIndex': 5,
                    },
                    'pasteType': 'PASTE_FORMAT',
                    'pasteOrientation': 'NORMAL'
                },

            }
        ],
        'includeSpreadsheetInResponse': False
    }

    service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=batch_update_request_body).execute()

    total_rows = len(rows)

    # write data 寫入E欄，寫入每一個產品的總剩餘存貨價值
    update_range = f"E1:E{total_rows}"
    # index 0 = E1的資料, index1 = E2資料...
    values = [['Values']]

    # 計算每一個產品的存貨價值 並加入要寫入的資料集
    for rowIndex in range(1, total_rows):
        row = rows[rowIndex]
        value = float(row[1]) * float(row[2])
        values.append([value])

    body = {
        'values': values
    }

    service.spreadsheets().values().update(spreadsheetId=sheet_id, range=update_range,
                                           valueInputOption='USER_ENTERED', body=body).execute()

except HttpError as err:
    print(err)
