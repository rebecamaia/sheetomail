from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from string import Template
import io
import shutil

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of the spreadsheet.
SPREADSHEET_ID = '13X9WTSFJs4G5Wc6O1cA5wRV6_qNIz83NwJxngs3ckzc'
RANGE_NAME = 'A:D'

def get_creds():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

def get_values(creds):
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        return values

def read_template(filename):
    with io.open(filename,  mode="r", encoding="utf-8") as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

def gen_body(values, message_template, row):
    message_template = read_template('message.txt')
    message = message_template.substitute(PERSON_NAME=row[1], INTERVIEW_DATE=row[3])
    return message

def send_emails(values, message_template):
    #0 timestamp
    #1 name
    #2 email
    #3 date
    for row in values[1:]:
        message = gen_body(values, message_template, row)
        print(message)

def main():
    creds = get_creds()
    values = get_values(creds)
    message_template = read_template('message.txt')
    send_emails(values, message_template)

if __name__ == '__main__':
    main()