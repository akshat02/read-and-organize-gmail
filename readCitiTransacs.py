import pandas as pd
import authorize as au

def get_messages(service):
    citiTransacsLabelID = 'Label_3370579833892165018'
    citiTransacsSearchQuery = 'from:CitiAlert.India@citicorp.com'
    user = 'me'

    results = service.users().messages().list(userId = user, labelIds = citiTransacsLabelID, maxResults = 5).execute()
    # results = service.users().messages().list(userId = 'me', q = citiTransacsSearchQuery).execute()
    messages = results.get('messages', [])

    return messages

def main():

    service = au.Authorization().authorize()

    messages = get_messages(service)

    data = []

    if not messages:
        print('No messages found')
    else:
        for message in messages:
            message_content = service.users().messages().get(userId = 'me', id = message['id']).execute()
            data.append([message_content['id'], message_content['payload']['headers'][7]['value'], message_content['payload']['headers'][26]['value'], message_content['snippet']])

    mails_df = pd.DataFrame(data, columns = ['Message Id', 'From', 'Subject', 'Snippet'])

    fileName = 'Citi_Gmail.xlsx'
    mails_df.to_excel(fileName)

    print('File Ready')

if __name__ == '__main__':
    main()
