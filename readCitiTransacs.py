import pandas as pd
import authorize as au
import re
from datetime import datetime
from multiprocessing import Process


def get_messages(service):
    citiTransacsLabelID = 'Label_3370579833892165018'
    citiTransacsSearchQuery = 'from:CitiAlert.India@citicorp.com'
    user = 'me'

    results = service.users().messages().list(userId=user, labelIds=citiTransacsLabelID, maxResults=5).execute()
    # results = service.users().messages().list(userId = 'me', q = citiTransacsSearchQuery).execute()
    messages = results.get('messages', [])

    return messages


def write_to_csv(df, filename):
    df.to_csv(filename + '.csv')


def write_to_excel(df, filename):
    df.to_excel(filename + '.xlsx')


def main():
    service = au.Authorization().authorize()
    messages = get_messages(service)

    data = []

    if not messages:
        print('No messages found')
    else:
        for message in messages:
            message_content = service.users().messages().get(userId='me', id=message['id']).execute()

            snippet = message_content['snippet']

            cost_match = re.search(r'\d+[.]\d+', snippet)
            if cost_match:
                cost = cost_match.group()
                cost = float(cost)
            else:
                cost = 'could not find'

            place_match = re.search(r'at\s[\w\s\/\,\*.]+\sLimit', snippet)
            if place_match:
                place = place_match.group()[3:-7]
            else:
                place = 'could not find'

            datetime_string = message_content['payload']['headers'][22]['value']
            dateTime = datetime.strptime(datetime_string[:-18], '%a, %d %b %Y %H:%M:%S')

            data.append([dateTime, cost, place, snippet])

    mails_df = pd.DataFrame(data, columns=['Date', 'Amount', 'Place', 'Snippet'])

    fileName = 'Citi_Gmail'

    p1 = Process(target=write_to_csv, args=(mails_df, fileName))
    p2 = Process(target=write_to_excel, args=(mails_df, fileName))

    p1.start()
    p2.start()

    p1.join()
    p2.join()

    print('File Ready')


if __name__ == '__main__':
    main()
