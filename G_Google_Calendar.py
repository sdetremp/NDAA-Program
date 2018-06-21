"""This module connects to the NDAA Calendar on alumaced@nd.edu.
   The Module is called through the main program and is not directly executed"""

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
SCOPES = 'https://www.googleapis.com/auth/calendar'
# If client secret file path is changed this will line will have to be updated
CLIENT_SECRET_FILE = r'G:\Team Drives\Alumni All Staff\Hesburgh Lecture Series\_HLS Code - DO NOT DELETE OR MOVE\_credentials\calendar_secret.json'
APPLICATION_NAME = 'Google Calendar API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    # If credentials file path is changed this will line will have to be updated
    credential_dir = r'G:\Team Drives\Alumni All Staff\Hesburgh Lecture Series\_HLS Code - DO NOT DELETE OR MOVE\_credentials'
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

    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    print('Getting the upcoming 10 events')
    eventsResult = service.events().list(
        calendarId='primary', timeMin=now, maxResults=10, singleEvents=True,
        orderBy='startTime').execute()
    events = eventsResult.get('items', [])

    if not events:
        print('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        print(start, event['summary'])


def create_event(ffirst, flast, clubname, date):
    # Parameters are for first and last name of Lecturer, name of the club, and date of the lecture
    """Creates an event on almaced calendar
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    event = {
        'summary': 'HLS - ' + ffirst[0] + '. ' + flast + ' (' + clubname + ')',
        'start': {
            'date': date
        },
        'end': {
            'date': date
        },
        'transparency': 'transparent'
    }
    # Calendar Id is below, only will need to be updated if major changes to Alumaced Calendar
    event = service.events().insert(calendarId='nd.edu_6i4u14dcleeb34obn4k9v01uds@group.calendar.google.com',
                                    body=event).execute()
    print('Event created: %s' % (event.get('htmlLink')))


if __name__ == '__main__':
    print('Google Calendar')
