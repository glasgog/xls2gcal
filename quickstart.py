
from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import datetime

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
#SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
SCOPES = 'https://www.googleapis.com/auth/calendar' #NOTE: scope modified because i want to write too
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Calendar API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():
    """Shows basic usage of the Google Calendar API.

    Creates a Google Calendar API service object and outputs a list of the next
    10 events on the user's calendar.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    CALENDAR = "test"   #scelgo il nome del calendario che mi interessa. NOTA: case sensitive
    calendar_id = None  #
    print('Getting the calendar list')
    page_token = None
    while True:
        calendar_list = service.calendarList().list(pageToken=page_token).execute()
        for calendar_list_entry in calendar_list['items']:
            print (calendar_list_entry['summary'], ":" , calendar_list_entry['id'])
            if calendar_list_entry['summary'] == CALENDAR:
                calendar_id = calendar_list_entry['id']
        page_token = calendar_list.get('nextPageToken')
        if not page_token:
            break

    if calendar_id == None:
        calendar_id = "primary" #l'id del calendario principale
    print (calendar_id)
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    print('Getting the upcoming 10 events')
    eventsResult = service.events().list(
        calendarId=calendar_id, timeMin=now, maxResults=10, singleEvents=True,
        orderBy='startTime').execute()
    events = eventsResult.get('items', [])

    if not events:
        print('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        print(start, event['summary'])

    #crea l'evento da aggiungere al calendario.
    #fare attenzione a dateTime, che deve essere utc.
    #nota che l'offset e' +01:00, mentre quando si e' in DST e' +02:00
    event = {
    'summary': 'Test Event',
    'location': 'Test location',
    'description': 'Test description',
    'start': {
    'dateTime': '2017-05-08T09:00:00+02:00',
    'timeZone': 'Europe/Rome',
    },
    'end': {
    'dateTime': '2017-05-08T09:30:00+02:00',
    'timeZone': 'Europe/Rome',
    },
    'reminders': {
    'useDefault': False,
    'overrides': [
    {'method': 'popup', 'minutes': 10},
    ],
    },
    }

    event = service.events().insert(calendarId=calendar_id, body=event).execute()
    print ('Event created: %s' % (event.get('htmlLink')))

    # created_event = service.events().quickAdd(
    # calendarId='calendar_id',
    # text='Appointment at Somewhere on May 8th 10am-10:25am').execute()

    # print (created_event['id'])

if __name__ == '__main__':
    main()