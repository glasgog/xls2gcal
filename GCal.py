"""
This file is part of xls2gcal
xls2gcal get a work shift from an excel file and add it to a given google calendar
Copyright (C) 2017  Ilario Digiacomo

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
"""

from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from datetime import datetime
import dateutil.parser

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modify these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
# SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
# NOTE: scope modified because i want to write too
SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'xls2gcal'


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
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


class GCal:
    # scelgo il nome del calendario che mi interessa (NOTA: case ensitive)
    # da rimuovere quando verra' implementata la possibilita' di creare un
    # nuovo calendario
    CALENDAR = "test"
    MAX_RESULTS = 10  # 365

    credentials = None
    service = None
    calendar_id = None

    def __init__(self, calendar_name=CALENDAR):
        """Shows basic usage of the Google Calendar API.

        Creates a Google Calendar API service object and outputs a list of the next
        MAX_RESULTS events on the user's calendar.
        """
        self.credentials = get_credentials()
        http = self.credentials.authorize(httplib2.Http())
        self.service = discovery.build('calendar', 'v3', http=http)

        # cerco l'id relativo al nome del calendario che mi interessa
        print('Getting the calendar list (name and id)...')
        page_token = None
        while True:
            calendar_list = self.service.calendarList().list(pageToken=page_token).execute()
            for calendar_list_entry in calendar_list['items']:
                print("-", calendar_list_entry['summary'],
                      "->", calendar_list_entry['id'])
                # Cerco l'id di calendar_name
                if calendar_list_entry['summary'] == calendar_name:
                    self.calendar_id = calendar_list_entry['id']
            page_token = calendar_list.get('nextPageToken')
            if not page_token:
                break
        # se non trovo il calendario che mi interessa, seleziono
        # l'id del calendario principale
        if self.calendar_id == None:
            self.calendar_id = "primary"
        print("Using id:", self.calendar_id)
        print()

        now = datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        print('Getting the upcoming', self.MAX_RESULTS, 'events...')
        eventsResult = self.service.events().list(
            calendarId=self.calendar_id, timeMin=now, maxResults=self.MAX_RESULTS, singleEvents=True,
            orderBy='startTime').execute()
        events = eventsResult.get('items', [])

        if not events:
            print('No upcoming events found.')
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print("-", start, event['summary'])
        print()

    def add_event(self, name="Work", start="2020-05-08T09:00:00+02:00", end="2020-05-08T09:30:00+02:00"):
        """
        Crea l'evento da aggiungere al calendario.
        E' preferibile che gli argomenti siano di tipo datatime:
        fare in ogni caso attenzione che siano utc o comunque correttamente localizzati
        (nota che l'offset e' +01:00, mentre quando si e' in DST e' +02:00)
        """

        # le formatto correttamente se non lo sono gia'
        if type(start) == datetime and type(end) == datetime:
            start = start.isoformat('T')
            end = end.isoformat('T')

        event = {
            'summary': name,
            'location': '',
            'description': '',
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

        event = self.service.events().insert(
            calendarId=self.calendar_id, body=event).execute()
        print('Event created: %s' % (event.get('htmlLink')))
        print()

    def event_on_date(self, date):
        """ Check if in a given date there is already an event"""
        now = datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        print('event_on_date: Getting the upcoming',
              self.MAX_RESULTS, 'events...')
        eventsResult = self.service.events().list(
            calendarId=self.calendar_id, timeMin=now, maxResults=self.MAX_RESULTS, singleEvents=True,
            orderBy='startTime').execute()  # NOTE: only the first 10 events are getted
        events = eventsResult.get('items', [])

        if not events:
            print('- No upcoming events found.')
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            # print("-", start, event['summary'])

            # convert from google api format
            start_dt = dateutil.parser.parse(start)
            if start_dt.date() == date.date():
                print("- Event found:", start, event['summary'])
                return True
        return False

    def update_event(self, date, new_name="Work", new_start="2017-05-08T09:00:00+02:00", new_end="2017-05-08T09:30:00+02:00"):
        # se avessi a disposizione l'id potrei evitare di rieffettuare la
        # ricerca
        # le formatto correttamente se non lo sono gia'
        if type(new_start) == datetime and type(new_end) == datetime:
            new_start = new_start.isoformat('T')
            new_end = new_end.isoformat('T')

        now = datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        print('update_event: Getting the upcoming',
              self.MAX_RESULTS, 'events...')
        eventsResult = self.service.events().list(
            calendarId=self.calendar_id, timeMin=now, maxResults=self.MAX_RESULTS, singleEvents=True,
            orderBy='startTime').execute()  # NOTE: only the first 10 events are getted
        events = eventsResult.get('items', [])

        if not events:
            print('- No upcoming events found.')
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            # print("-", start, event['summary'])

            # convert from google api format
            start_dt = dateutil.parser.parse(start)
            if start_dt.date() == date.date():
                print("- Event found:", start, event['summary'])
                event["summary"] = new_name
                event["start"] = {'dateTime': new_start,
                                  'timeZone': 'Europe/Rome', }
                event["end"] = {'dateTime': new_end,
                                'timeZone': 'Europe/Rome', }
                updated_event = self.service.events().update(calendarId=self.calendar_id,
                                                             eventId=event['id'], body=event).execute()
                # Print the updated date.
                print("  Event updated:", updated_event['updated'])
                return True
        return False

    def delete_event(self, date):

        now = datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        print('delete_event: Getting the upcoming',
              self.MAX_RESULTS, 'events...')
        eventsResult = self.service.events().list(
            calendarId=self.calendar_id, timeMin=now, maxResults=self.MAX_RESULTS, singleEvents=True,
            orderBy='startTime').execute()  # NOTE: only the first 10 events are getted
        events = eventsResult.get('items', [])

        if not events:
            print('- No upcoming events found.')
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            # print("-", start, event['summary'])

            # convert from google api format
            start_dt = dateutil.parser.parse(start)
            if start_dt.date() == date.date():
                print("- Event found:", start, event['summary'])
                self.service.events().delete(calendarId=self.calendar_id,
                                             eventId=event['id']).execute()
                print("  Event deleted.")
                return True
        return False


def main():
    # TOFIX when class is modified
    print("Main is stil not implemented.")

if __name__ == '__main__':
    main()
