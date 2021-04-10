from __future__ import print_function
import pickle
import os.path
import re
import math
import datetime
import base64
import email
import html2text
from pprint import pprint
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# User defined configuration
LABEL_NAME = 'billetin'
LABEL_ID = 'Label_5982538524887334744'
SHEET = '1RTaiXlVJYcIMhFlY1YniACN7XZgUMcWy9__MgMSPhFY'

# If modifying these scopes, delete the file token.pickle.
SCOPES = [
        'https://www.googleapis.com/auth/gmail.readonly',
        'https://www.googleapis.com/auth/spreadsheets'
        ]
#PATTERN = r'en el establecimiento (.*) por $ (.*) , el (.*)'
MAX=150
SQRT=int(math.ceil(math.sqrt(MAX)))

def multipija(mime_msg):
    messageMainType = mime_msg.get_content_maintype()
    if messageMainType == 'multipart':
        for part in mime_msg.get_payload():
            if part.get_content_maintype() == 'text':
                return part.get_payload()
        return ""
    elif messageMainType == 'text':
        return mime_msg.get_payload()

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

    # Uncomment this block to get the ID from the chosen label on Gmail.
    # results = service.users().labels().list(userId='me').execute()
    # labels = results.get('labels', [])
    # if not labels:
    #     print('No labels found.')
    # else:
    #     print('Labels:')
    #     for label in labels:
    #         if label['name'] == LABEL_NAME:
    #             print(label)
    #     exit(0)

    results = service.users().messages().list(
            userId='me',
            labelIds=[LABEL_ID],
            maxResults=MAX,
            q='nuevos consumos de sus tarjetas visa'
            ).execute()
    messageIdsList = results.get('messages', [])

    expenses = [('Donde', 'Cuanto', 'Cuando', '$/u$s', 'Cuotas', 'Tarjeta')]
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
                message = service.users().messages().get(userId='me', id=messageId, format='raw').execute()
                ascii_raw_message = message['raw'].encode('ASCII')
                decoded_bytes_or_str = base64.urlsafe_b64decode(ascii_raw_message)
                email_message = email.message_from_bytes(decoded_bytes_or_str)
                html = multipija(email_message)
                text = html2text.html2text(html)
                clean_text = text.replace('*', '')
                parts = re.findall(r'.*en\sel\sestablecimiento\s(.*)\spor\s([^ ]*)\s([^ ]*)(.*)\s,\sel\s(.*)\sa\slas\s.*Tarjeta de [^/]*/(..).*', clean_text, flags=re.DOTALL)[0]

                cuotas = parts[3]
                if 'cuotas' in cuotas:
                    cuotas = int(re.findall(r'.en (.*) cuotas.', cuotas)[0])
                else:
                    cuotas = 1
                parts = (parts[0], float(parts[2])/cuotas, parts[4], parts[1], cuotas, parts[5])
                expenses.append(parts)

    print()
    pprint(expenses)

    sheetyService = build('sheets', 'v4', credentials=creds)

    results = sheetyService.spreadsheets().values().update(
            spreadsheetId=SHEET,
            valueInputOption='RAW',
            range='A1:F' + str(MAX+1),
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
