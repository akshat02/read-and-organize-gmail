from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

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

    citiTransacsLabelID = 'Label_3370579833892165018'
    citiTransacsSearchQuery = 'from:CitiAlert.India@citicorp.com'

    results = service.users().messages().list(userId = 'me', labelIds = citiTransacsLabelID, maxResults = 5).execute()
    # results = service.users().messages().list(userId = 'me', q = citiTransacsSearchQuery).execute()
    messages = results.get('messages', [])

    data = []

    if not messages:
        print('No messages found')
    else:
        for message in messages:
            message_content = service.users().messages().get(userId = 'me', id = message['id']).execute()
            # print(message_content['snippet'])
            data.append([message_content['id'], message_content['payload']['headers'][23]['value'], message_content['snippet']])

    df = pd.DataFrame(data, columns = ['Message Id', 'From', 'Snippet'])

    fileName = 'Citi_Gmail.xlsx'
    df.to_excel(fileName)
    print('File Ready')

if __name__ == '__main__':
    main()
