from __future__ import print_function
import pickle
import os.path
import re
import math
import datetime
from pprint import pprint
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# If modifying these scopes, delete the file token.pickle.
SCOPES = [
        'https://www.googleapis.com/auth/gmail.readonly',
        'https://www.googleapis.com/auth/spreadsheets'
        ]
LABEL_NAME = 'billetin'
LABEL_ID = 'Label_5982538524887334744'
PATTERN = r'en el establecimiento (.*) por $ (.*) , el (.*)'
SHEET = '1RTaiXlVJYcIMhFlY1YniACN7XZgUMcWy9__MgMSPhFY'
MAX=100
SQRT=int(math.ceil(math.sqrt(MAX)))

def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
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

    service = build('gmail', 'v1', credentials=creds)

    # Call the Gmail API
    # results = service.users().labels().list(userId='me').execute()
    # labels = results.get('labels', [])

    # if not labels:
    #     print('No labels found.')
    # else:
    #     print('Labels:')
    #     for label in labels:
    #         if label['name'] == LABEL_NAME:
    #             print(label)

    results = service.users().messages().list(
            userId='me',
            labelIds=[LABEL_ID],
            maxResults=MAX,
            q='nuevos consumos de sus tarjetas visa'
            ).execute()
    messageIdsList = results.get('messages', [])

    expenses = []
    if os.path.exists('expenses.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    else:
        if not messageIdsList:
            print('No messages found.')
        else:
            i=0
            for messageIds in messageIdsList:
                i+=1
                messageId = messageIds['id']
                print(u'██', end='', flush=True)
                if i % SQRT == 0:
                    print()
                message = service.users().messages().get(userId='me', id=messageId).execute()
                snippet = message['snippet']
                #print(snippet)
                parts = re.findall(r'.*en el establecimiento (.*) por ([^ ]*) ([^ ]*)(.*) , el (.*) a las .*', snippet)[0]
                #print(parts)

                cuotas = parts[3]
                if 'cuotas' in cuotas:
                    cuotas = re.findall(r'.en (.*) cuotas.', cuotas)[0]
                parts = (parts[0], float(parts[2])/cuotas, parts[4], parts[1], cuotas)
                expenses.append(parts)

    print()
    pprint(expenses)

    sheetyService = build('sheets', 'v4', credentials=creds)

    results = sheetyService.spreadsheets().values().update(
            spreadsheetId=SHEET,
            valueInputOption='RAW',
            range='A1:F100',
            body=dict(
                    majorDimension='ROWS',
                    values=expenses
                )
            ).execute()

    print()
    print('saved')
    pprint(results)

if __name__ == '__main__':
    main()
