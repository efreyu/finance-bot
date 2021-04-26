from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import json

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']


def get_setting(name):
    with open('parseSettings.json') as json_file:
        data = json.load(json_file)
        if not data[name]:
            print("%s is not valid!" % name)
        return data[name]


def parseSheet():
    spreadsheet_id = get_setting('spreadsheetId')
    findDiapason = get_setting('findDiapason')
    listPage = get_setting('listPage')
    # parse rules:
    # - data
    # - category
    # - calc income
    # - calc currencies
    # - print to page
    data = {}
    for page in get_setting('pages'):
        pageName = get_sheet_name_by_id(spreadsheet_id, page['id'])
        pageData = parseData(spreadsheet_id, pageName + "!" + findDiapason)
        for item in pageData:
            if not item[0] in data:
                data[item[0]] = {}
            if not item[2] in data[item[0]]:
                data[item[0]][item[2]] = float(item[1])
            else:
                data[item[0]][item[2]] = float(data[item[0]][item[2]]) + float(item[1])
    test = ""
        # data.append()


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


def main():
    parseSheet()


if __name__ == '__main__':
    main()
