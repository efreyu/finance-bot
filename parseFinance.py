from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import json
from datetime import datetime

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SCOPES_EDIT = ['https://www.googleapis.com/auth/spreadsheets']


def get_setting(name):
    with open('parseSettings.json') as json_file:
        data = json.load(json_file)
        if not data[name]:
            print("%s is not valid!" % name)
        return data[name]


def parseSheet():
    spreadsheet_id = get_setting('spreadsheetId')
    findDiapason = get_setting('findDiapason')
    listPage = get_sheet_name_by_id(spreadsheet_id, get_setting('listPage'))
    incomeCat = parseData(spreadsheet_id, listPage + "!G2")[0][0]
    incomeCats = parseData(spreadsheet_id, listPage + "!G2:G")
    incomeDate = parseData(spreadsheet_id, listPage + "!F2")[0][0]
    datetime_income = datetime.strptime(incomeDate, '%Y-%m-%d')

    data = {}
    for page in get_setting('pages'):
        pageName = get_sheet_name_by_id(spreadsheet_id, page['id'])
        pageData = parseData(spreadsheet_id, pageName + "!" + findDiapason)
        pageCur = parseData(spreadsheet_id, pageName + "!F2")[0][0]
        if not pageCur in data:
            data[pageCur] = {}
        startBalance = parseData(spreadsheet_id, pageName + "!F1")[0][0]
        if startBalance:
            startBalance = float(startBalance)
        if not datetime_income in data[pageCur]:
            data[pageCur][datetime_income] = {}
        if not incomeCat in data[pageCur][datetime_income]:
            data[pageCur][datetime_income][incomeCat] = startBalance
        else:
            data[pageCur][datetime_income][incomeCat] = float(data[pageCur][datetime_income][incomeCat]) + startBalance
        for item in pageData:
            datetime_key = datetime.strptime(item[0], '%Y-%m-%d')
            if not datetime_key in data[pageCur]:
                data[pageCur][datetime_key] = {}
            if not item[2] in data[pageCur][datetime_key]:
                data[pageCur][datetime_key][item[2]] = float(item[1])
            else:
                data[pageCur][datetime_key][item[2]] = float(data[pageCur][datetime_key][item[2]]) + float(item[1])
    result = {}
    for cur in data:
        if not cur in result:
            result[cur] = {}
        keylist = data[cur].keys()
        for key in sorted(keylist):
            row = data[cur][key]
            month = str(key.year) + "-" + str(key.month)
            if not month in result[cur]:
                result[cur][month] = {}
            cats = row.keys()
            for cat in sorted(cats):
                if not cat in result[cur][month]:
                    result[cur][month][cat] = row[cat]
                else:
                    result[cur][month][cat] += row[cat]
    return result


def parseData(spreadsheetId, findRange):
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheetId,
                                range=findRange).execute()
    values = result.get('values', [])
    return values


def get_sheet_name_by_id(spreadsheetId, id):
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    sheet_metadata = sheet.get(spreadsheetId=spreadsheetId).execute()
    sheets = sheet_metadata.get('sheets', '')
    return sheets[id].get("properties", {}).get("title", "Sheet1")


def create_write_service():
    cred = None

    if os.path.exists('token_write.pickle'):
        with open('token_write.pickle', 'rb') as token:
            cred = pickle.load(token)

    if not cred or not cred.valid:
        if cred and cred.expired and cred.refresh_token:
            cred.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES_EDIT)
            cred = flow.run_local_server()

        with open('token_write.pickle', 'wb') as token:
            pickle.dump(cred, token)

    try:
        service = build('sheets', 'v4', credentials=cred)
        print('sheets', 'service created successfully')
        return service
    except Exception as e:
        print(e)
        return None


def convert_to_list(data):
    result = list()
    spreadsheet_id = get_setting('spreadsheetId')
    listPage = get_sheet_name_by_id(spreadsheet_id, get_setting('listPage'))
    allCats = parseData(spreadsheet_id, listPage + "!A2:A")
    dataWithCats = {}
    firstRow = list()
    for cur in data:
        firstRow.append(str(cur))
        if len(firstRow) == 1:
            for item in allCats:
                dataWithCats[item[0]] = [str(item[0])]
        else:
            for item in allCats:
                dataWithCats[item[0]].append(str(item[0]))
        for month in sorted(data[cur].keys(), reverse=True):
            firstRow.append(str(month))
        for month in sorted(data[cur].keys(), reverse=True):
            count = 0
            for cat in data[cur][month].keys():
                if cat in dataWithCats.keys():
                    dataWithCats[cat].append(data[cur][month][cat])
                    if count < len(dataWithCats[cat]):
                        count = len(dataWithCats[cat])
            for cat in dataWithCats.keys():
                if count > len(dataWithCats[cat]):
                    for i in range(count - len(dataWithCats[cat])):
                        dataWithCats[cat].append('')


    result.append(firstRow)
    for cat in dataWithCats:
        result.append(dataWithCats[cat])
    return result

def main():
    data = parseSheet()
    data = convert_to_list(data)
    # dataFrame = parseData(get_setting('spreadsheetId'), get_setting('editDiapason'))
    service = create_write_service()
    response_date = service.spreadsheets().values().update(
        spreadsheetId=get_setting('spreadsheetId'),
        valueInputOption='RAW',
        range=get_setting('editDiapason'),
        body=dict(
            majorDimension='ROWS',
            values=data
        )
    ).execute()
    print('Sheet successfully Updated')


if __name__ == '__main__':
    main()
