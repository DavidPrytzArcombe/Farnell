# Bankgiro reader

import csv
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import base64
from email.message import EmailMessage
from datetime import date
from dateutil.relativedelta import relativedelta


# Scopes that define what the app should be allowed to do. If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/gmail.send"]

# For reading and editing Google sheets. The ID and range of a sample spreadsheet.
spreadsheetID = "1TkXa5Vo6YjOQMRKHpBH-HKvMNLktoyd7-nWEad61Ffc"


def readBankGiro():
    with open('bankgirobetalningar.csv',mode='r') as file:
        csvFile = csv.reader(file)
        for _ in range(14):
            next(csvFile)
        payments = []
        for line in csvFile:
            print(' ')
            firstRow = line[0].split(';')
            if len(firstRow) != 5 or len(firstRow[2]) == 0:
                print('Skum entry:',line)
            else:
                invoiceNumer = firstRow[2]
                name = firstRow[1]
                amountOwed = firstRow[4].replace(' ','')+'.'+line[1]
                payment = [invoiceNumer,name,amountOwed]
                payments.append(payment)
                print(payment)


def readGoogleSheet():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()
        result = (sheet.values().get(spreadsheetId=spreadsheetID, range='Invoices!A2:A100').execute())
        values = result.get("values", [])
        print(values)
    except HttpError as err:
        print(err)



def main():
    #readBankGiro()
    readGoogleSheet()


if __name__ == '__main__':
    main()