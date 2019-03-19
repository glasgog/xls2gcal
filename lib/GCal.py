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
                                   APPLICATION_NAME+'.json')

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
    # TODO: creazione nuovo calendario se non esistente
    CALENDAR = "test" # nome del calendario su cui aggiungere gli eventi (case sensitive)
    MAX_RESULTS = 60  # numero di eventi da considerare

    credentials = None
    service = None
    calendar_id = None
    events_to_read = None
    events = None

    def __init__(self, calendar_name=CALENDAR, events_to_read=MAX_RESULTS):
        """Shows basic usage of the Google Calendar API.

        Creates a Google Calendar API service object and outputs a list of the next
        MAX_RESULTS events on the user's calendar.
        """
        self.credentials = get_credentials()
        http = self.credentials.authorize(httplib2.Http())
        self.service = discovery.build('calendar', 'v3', http=http)

        print()
        # cerco l'id relativo al nome del calendario che mi interessa
        print('Getting the calendar list (name -> id)...')
        page_token = None
        while True:
            calendar_list = self.service.calendarList().list(pageToken=page_token).execute()
            for calendar_list_entry in calendar_list['items']:
                # print("-", calendar_list_entry['summary'], "->", calendar_list_entry['id'])
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
        print("Use id", self.calendar_id)
        # imposto il numero di eventi da leggere. Nota che giorni=N implica eventi<=N
        # in un intervallo di N giorni non possono esserci piu' di N eventi.
        # al piu' verranno letti eventi oltre il giorno di interesse
        self.events_to_read = events_to_read
        print()
        self.events = self.get_events()


    def get_events(self):
        """ Get event list """
        now = datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        print('Getting the upcoming', self.events_to_read, 'events...')
        eventsResult = self.service.events().list(
            calendarId=self.calendar_id, timeMin=now, maxResults=self.events_to_read, singleEvents=True,
            orderBy='startTime').execute()
        events = eventsResult.get('items', [])
        return events

    def event_on_date(self, date, get_event=False):
        """
        Check if in a given date there is already an event
            INPUT: date and get_event. If the latter is True
                the event object is returned if found
            OUTPUT: False if not found, True or event object
                in the other case
        """
        events = self.events
        if not events:
            print('  No upcoming events found.')
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            # print("-", start, event['summary'])
            # convert from google api format
            start_dt = dateutil.parser.parse(start)
            if start_dt.date() == date.date():
                print("  Event found:", start, event['summary'])
                if get_event:
                    return event
                return True
        return False

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
        # print('Event created: %s' % (event.get('htmlLink')))
        print('  Event created:', start, event['summary'])

    def update_event(self, date, new_name="Work", new_start="2017-05-08T09:00:00+02:00", new_end="2017-05-08T09:30:00+02:00"):
        """
        If there is already an event on a given date, it is updated.
            INPUT: date, event new name and datetime
            OUTPUT: True if event is updated, False if no event on that date
        """
        # TODO: evitare di effettuare l'update se l'evento presente e' gia' identico a quello da inserire
        if type(new_start) == datetime and type(new_end) == datetime:
            new_start = new_start.isoformat('T')
            new_end = new_end.isoformat('T')

        event = self.event_on_date(date, get_event=True)
        if event:
            event["summary"] = new_name
            event["start"] = {'dateTime': new_start,
                              'timeZone': 'Europe/Rome', }
            event["end"] = {'dateTime': new_end,
                            'timeZone': 'Europe/Rome', }
            updated_event = self.service.events().update(calendarId=self.calendar_id,
                                                         eventId=event['id'], body=event).execute()
            # Print the updated date.
            print("  Event updated:", new_start, updated_event['summary'])
            print("  last update:", updated_event['updated'])
            return True
        return False

    def delete_event(self, date):
        event = self.event_on_date(date, get_event=True)
        if event:
            self.service.events().delete(calendarId=self.calendar_id,
                                         eventId=event['id']).execute()
            print("  Event deleted.")
            return True
        return False

    def print_events(self):
        events = self.events
        if not events:
            print('No upcoming events found.')
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print("-", start, event['summary'])


def main():
    # TOFIX when class is modified
    print("Main is stil not implemented.")

if __name__ == '__main__':
    main()
