
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


class GCal:

    credentials = None
    service = None
    #scelgo il nome del calendario che mi interessa (NOTA: case sensitive)
    #e cerco il suo id
    CALENDAR = "test"   
    calendar_id = None

    def __init__(self):
        """Shows basic usage of the Google Calendar API.

        Creates a Google Calendar API service object and outputs a list of the next
        10 events on the user's calendar.
        """
        self.credentials = get_credentials()
        http = self.credentials.authorize(httplib2.Http())
        self.service = discovery.build('calendar', 'v3', http=http)

        #cerco l'id relativo al nome del calendario che mi interessa
        print('Getting the calendar list (name and id):')
        page_token = None
        while True:
            calendar_list = self.service.calendarList().list(pageToken=page_token).execute()
            for calendar_list_entry in calendar_list['items']:
                print ("-", calendar_list_entry['summary'], "->" , calendar_list_entry['id'])
                if calendar_list_entry['summary'] == self.CALENDAR:
                    self.calendar_id = calendar_list_entry['id']
            page_token = calendar_list.get('nextPageToken')
            if not page_token:
                break

        #se non trovo il calendario che mi interessa, seleziono
        #l'id del calendario principale
        if self.calendar_id == None:
            self.calendar_id = "primary" #
        print ("Using id:", self.calendar_id)

        now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
        print('Getting the upcoming 10 events')
        eventsResult = self.service.events().list(
            calendarId=self.calendar_id, timeMin=now, maxResults=10, singleEvents=True,
            orderBy='startTime').execute()
        events = eventsResult.get('items', [])

        if not events:
            print('No upcoming events found.')
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print("-", start, event['summary'])

    def add_event(self, start="2017-05-08T09:00:00+02:00", end="2017-05-08T09:30:00+02:00"):
        #crea l'evento da aggiungere al calendario.
        #fare attenzione a dateTime, che deve essere utc.
        #nota che l'offset e' +01:00, mentre quando si e' in DST e' +02:00
        event = {
        'summary': 'Test Event',
        'location': 'Test location',
        'description': 'Test description',
        'start': {
        'dateTime': start,
        'timeZone': 'Europe/Rome',
        },
        'end': {
        'dateTime': end,
        'timeZone': 'Europe/Rome',
        },
        'reminders': {
        'useDefault': False,
        'overrides': [
        {'method': 'popup', 'minutes': 10},
        ],
        },
        }

        event = self.service.events().insert(calendarId=self.calendar_id, body=event).execute()
        print ('Event created: %s' % (event.get('htmlLink')))

    def event_on_day(self, day):
        """ Check if in a given day (int) there is already an event """
        now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
        print('Getting the upcoming 10 events')
        eventsResult = self.service.events().list(
            calendarId=self.calendar_id, timeMin=now, maxResults=10, singleEvents=True,
            orderBy='startTime').execute()
        events = eventsResult.get('items', [])
        
        if not events:
            print('No upcoming events found.')
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print("-", start, event['summary'])

            #convert from google api format
            import dateutil.parser #HARDCODED
            dt = dateutil.parser.parse(start)
            if dt.day == day:
                print("Event found")
                return True
        return False



def main():
    #TOFIX when class is modified
    c = GCal()
    c.add_event()

if __name__ == '__main__':
    main()