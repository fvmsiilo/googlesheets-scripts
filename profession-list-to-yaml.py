from __future__ import print_function
import codecs
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

SPREADSHEET_ID = '15cpxPUJ5mtMuXO43XDlT1AA6kcOCEna0rs3Jk769z8A'
RANGE_DATA = 'professions!A2:D79'
OUTPUT_FILE = 'profession_tree.yaml'

def main():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
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

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_DATA).execute()
    rows = result.get('values', [])
    disallowCategories = ['5']

    if not rows:
        print('No data found.')
    else:

        lines = ['profession-data.categories:']
        printLastLine(lines)

        cursorAtProfession = False
        for row in rows:
            if row[0] == 'c':
                if cursorAtProfession is True: 
                    cursorAtProfession = False
                lines.append(u'  - id: {0} \t # --------- {1} --------- #'.format(row[1], row[3]))
                printLastLine(lines)
                if row[1] in disallowCategories:
                    disallow = 'true'
                else:
                    disallow = 'false'
                lines.append(u'    disallow-training: {0}'.format(disallow))
                printLastLine(lines)

            if row[0] == 'p':
                if cursorAtProfession is False:
                    lines.append(u'    professions:')
                    printLastLine(lines)
                    cursorAtProfession = True
                lines.append(u'    - id: {0} \t # {1}'.format(row[2], row[3]))
                printLastLine(lines)
        
        if (len(lines) > 2) is True:
            with codecs.open(OUTPUT_FILE, 'w', 'utf-8') as f:
                f.write('\n'.join(lines))
                f.close()

def printLastLine(array):
    if (len(array) > 0) is True: 
        print(array[len(array) - 1])

if __name__ == '__main__':
    main()
