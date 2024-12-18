# Farnell Google API functions

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
spreadsheetID = "1GQ7sYGzivvrQ2fROXODhQSH10eJmewx7gSCMAzSKCAg"


def readOldInvoiceNumbers():
    creds = getCredentials()
    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()
        result = (sheet.values().get(spreadsheetId=spreadsheetID, range='Invoices!A2:A10000').execute())
        values = result.get("values", [])
        return values
    except HttpError as err:
        print(err)


def checkPOnumbers(POnumber):
    creds = getCredentials()
    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()
        result = (sheet.values().get(spreadsheetId=spreadsheetID, range='PO numbers!A2:A1000000').execute())
        values = result.get("values", [])
        if [POnumber] in values:
            print('You have already sent invoices for this order.')
            return True
        else:
            return False
    except HttpError as err:
        print(err)


def addPOnumber(POnumber):
    creds = getCredentials()
    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()
        usedRows = sheet.values().get(spreadsheetId=spreadsheetID, range='PO numbers!A1:A' + '1000000').execute()
        inputRow = str(len(usedRows.get('values',[]))+1)
        writeRange = 'PO numbers!A' + inputRow
        sheet.values().update(spreadsheetId=spreadsheetID, range=writeRange,valueInputOption='USER_ENTERED',body={'values':[[POnumber]]}).execute()
    except HttpError as err:
        print(err)


def readCurrentEmailAddresses():
    creds = getCredentials()
    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()
        result = (sheet.values().get(spreadsheetId=spreadsheetID, range='Email addresses!A2:B10000').execute())
        values = result.get("values", [])
        return values
    except HttpError as err:
        print(err)


def addInvoiceToSheet(entry):
    creds = getCredentials()
    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()
        usedRows = (sheet.values().get(spreadsheetId=spreadsheetID, range= 'Invoices!A1:A' + '100000').execute())
        inputRow = str(len(usedRows.get('values',[]))+1)
        columnLetters = alphabet()
        writeRange = 'Invoices!A' + inputRow + ':A' + str(columnLetters[len(entry[0])-1]) + inputRow
        sheet.values().update(spreadsheetId=spreadsheetID, range=writeRange,valueInputOption='USER_ENTERED',body={'values':entry}).execute()
    except HttpError as err:
        print(err)


def createSingleEntryForSheet(order,invoiceNumber):
    entry = []
    entry.append(invoiceNumber)
    entry.append(order[0][0] + ' ' + order[0][1])   # Name
    entry.append(order[1][0][0])                    # Date when ETA placed the Farnell order
    entry.append('%.2f' % float(order[2]))          # Total amount owed
    for item in order[1]:
        itemNumber = item[1]
        unitQuantity = '%.0f' % float(item[2])
        unitPrice = '%.5f' % float(item[3][3:])
        unitPriceTimesQuantity = '%.2f' % float(item[4][3:])
        entry.append('Produktnummer: ' + itemNumber + ', Antal: ' + unitQuantity + ', Á-pris: ' + unitPrice + ', Summa: ' + unitPriceTimesQuantity)
    return [entry]


def sendInvoice(name,address,invoiceNumber):
    creds = getCredentials()
    due_date = date.today() + relativedelta(months=+3)
    try:
        service = build("gmail", "v1", credentials=creds)
        message = EmailMessage()
        text = 'Hej ' + name + '.\n\nDu har fått en faktura för Farnellförmedling via ETA, se bifogat dokument.\nDen skall vara betald senast ' + str(due_date)
        text += '''\n\nVänligen kontakta oss om du har några frågor gällande din faktura.\n\nE-sektionens Teletekniska Avdelning 
            \nETA Chalmers \nRännvägen 4 \n41296 Göteborg \n\n031-20 78 60 \neta@eta.chalmers.se'''
        message.set_content(text)
        message["To"] = address
        message["From"] = "david.prytz.arcombe@eta.chalmers.se"
        message["Subject"] = "Faktura " + invoiceNumber + " ETA Farnellförmedling"

        attachment_filename = "ETA Faktura " + name + " " + invoiceNumber + ".pdf"
        maintype = 'application'; subtype = 'pdf'
        with open(attachment_filename, "rb") as fp:
            attachment_data = fp.read()
        message.add_attachment(attachment_data, maintype, subtype, filename=attachment_filename)

        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        create_message = {"raw": encoded_message}
        send_message = (service.users().messages().send(userId="me", body=create_message).execute())
        print(f'Message Id: {send_message["id"]}')
    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f"An error occurred: {error}")


def getCredentials():
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
    return creds


def alphabet():
    alphabet = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
    columnLetters = [alphabet]
    for letter in alphabet:
        for subletter in alphabet:
            columnLetters.append(letter + subletter)
    return columnLetters


def main():
    print('This is the Google API script.')


if __name__ == "__main__":
    main()