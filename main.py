from __future__ import print_function
import event_helper
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from os import path

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']


def main():
    """Creates Google Calander events from data in Excel sheets.
    Gets event name and date from event_helper script,
    Uses Google API to create calander events.
    """
    service = get_service()
    Excel_file_path = input("Enter the name of the excel file: ")
    while not path.isfile(Excel_file_path):
        print("The file name entered does not exist!")
        Excel_file_path = input("Enter the name of the excel file (make sure it is placed in the same folder): ")
    sheet_index = int(input("Which sheet number in the file do you want to scan [1, 2, 3 etc.]: ")) - 1
    key = input("Enter the name you want to search in the excel file: ")

    # Geting event data from the excel file
    events = event_helper.get_events(Excel_file_path, sheet_index, key)

    # Printing, creating and uploading events
    for eve in events:
        (eventName, eventDate) = eve
        event = {
            'summary': eventName,
            'start': {
                "date": eventDate,
                'timeZone': 'Asia/Dubai'
            },
            'end': {
                "date": eventDate,
                'timeZone': 'Asia/Dubai'
            },
            'reminders': {
                'useDefault': False,
                'overrides': [
                    {'method': 'popup', 'minutes': 6 * 60},         # 6 hours
                    {'method': 'popup', 'minutes': 1 * 24 * 60},    # 1 day
                    {'method': 'popup', 'minutes': 2 * 24 * 60},    # 2 days
                    {'method': 'popup', 'minutes': 3 * 24 * 60}     # 3 days
                ],
            },
            'colorId': '3'
        }

        event = service.events().insert(calendarId='primary', body=event).execute()
        print("\nEvent name = '", eventName, "', Event date = '", eventDate, "'", sep='')
        print("Event Link:", format(event.get('htmlLink')))

    if len(events) == 0:
        print("No events found in this sheet.\nTry changing value of 'sheet index'")

def get_service():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
            print("Token pickle found")
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists('credentials.json'):
                print("\nCreate a credientials.json file and copy it to this folder.")
                print("\nhttps://developers.google.com/calendar/quickstart/python")
                print("Go to this site, click the Enable Google Calander API button, and copy the downloaded credientials.json file to this folder")
                input("Press any key to end the program")
                exit()
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)
    return service


if __name__ == '__main__':
    main()
    input("Press any key to end the program")


# 41beb2e09877fadffac346651f3c78a4429b7433
# 4dd4c9fd9743d942dd70a93df42b6e4e
