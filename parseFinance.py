from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import json

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

def get_spread_sheet_id():
    with open('parseSettings.json') as json_file:
        data = json.load(json_file)
        if not data['spreadsheetId']:
            print("spreadsheetId is not valid!")
        return data['spreadsheetId']


def main():
    spreadsheet_id = get_spread_sheet_id()
    print('%s' % spreadsheet_id)
    # findRange = getFindRange()
    # print('%s' % (startRow))
    # dataStruct = getDataStruct()
    # parseSheet(spreadsheetId, findRange, dataStruct)


if __name__ == '__main__':
    main()
