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
RANGE_DATA = 'professions!A2:F80'

FILENAME = 'professions'
FILENAME_EXTENSION = '.properties'

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

    if not rows:
        print('No data found.')
    else:

        lines_en = ['# English']
        lines_nl = ['# Dutch']
        lines_de = ['# German']
        # lines_es = ['# Spanish']
        # lines_fr = ['# French']
        # lines_it = ['# Italian']
        # lines_pt = ['# Portuguese']
        # lines_ru = ['# Russian']

        lines_en.append('# New-style professions')
        lines_nl.append('# New-style professions')
        lines_de.append('# New-style professions')

        for row in rows:
            if row[0] == 'c':
                lines_en.append(u'category.{0}={1}'.format(row[1], row[3]))
                lines_nl.append(u'category.{0}={1}'.format(row[1], row[4]))
                lines_de.append(u'category.{0}={1}'.format(row[1], row[5]))
            if row[0] == 'p':
                lines_en.append(u'profession.{0}={1}'.format(row[2], row[3]))
                lines_nl.append(u'profession.{0}={1}'.format(row[2], row[4]))
                lines_de.append(u'profession.{0}={1}'.format(row[2], row[5]))
        
        lines_en.append('# end')
        lines_nl.append('# end')
        lines_de.append('# end')

        # Write EN properties file
        if (len(lines_en) > 2) is True:
            with codecs.open(FILENAME + FILENAME_EXTENSION, 'w', 'utf-8') as f:
                # f.write(u'\ufeff')
                f.write('\n'.join(lines_en))
                f.close()
        # Write NL properties file
        if (len(lines_nl) > 2) is True:
            with codecs.open(FILENAME + '_nl' + FILENAME_EXTENSION, 'w', 'utf-8') as f:
                # f.write(u'\ufeff')
                f.write('\n'.join(lines_nl))
                f.close()
        # Write DE properties file
        if (len(lines_de) > 2) is True:
            with codecs.open(FILENAME + '_de' + FILENAME_EXTENSION, 'w', 'utf-8') as f:
                # f.write(u'\ufeff')
                f.write('\n'.join(lines_de))
                f.close()

if __name__ == '__main__':
    main()
