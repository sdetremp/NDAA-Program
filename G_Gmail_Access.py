"""This module allows users to create drafts for the alumaced@nd.edu email account
   The Module is called through the main program and is not directly executed"""

from __future__ import print_function
import httplib2
import os

import apiclient
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import base64
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import mimetypes

try:
    import argparse

    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/gmail-python-quickstart.json
SCOPES = ['https://www.googleapis.com/auth/gmail.compose',
          'https://www.googleapis.com/auth/gmail.modify',
          'https://www.googleapis.com/auth/gmail.send']

# If client secret file path is changed this will line will have to be updated
CLIENT_SECRET_FILE = r'G:\Team Drives\Alumni All Staff\Hesburgh Lecture Series\_HLS Code - DO NOT DELETE OR MOVE\_credentials\client_secret.json'
APPLICATION_NAME = 'Gmail API Python Quickstart'


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
                                   'gmail-python-quickstart.json')

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


def main():
    """Shows basic usage of the Gmail API.

    Creates a Gmail API service object and outputs a list of label names
    of the user's Gmail account.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = apiclient.discovery.build('gmail', 'v1', http=http)

    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])

    if not labels:
        print('No labels found.')
    else:
        print('Labels:')
        for label in labels:
            print(label['name'])


def create_message_with_attachment(sender, to, cc, subject, message_text, attachments, verbose=True):
    """Create a message for an email.

    Args:
      sender: Email address of the sender.
      to: Email address of the receiver.
      cc: Anyone cc'd?
      subject: The subject of the email message.
      message_text: The text of the email message.
      attachments: The path to the file to be attached
      verbose: Is there a link or not.

    Returns:
      An object containing a base64url encoded email object.
    """

    message = MIMEMultipart('alternative')
    message['to'] = to
    message['cc'] = cc
    message['from'] = sender
    message['subject'] = subject

    if verbose:
        msg1 = MIMEText(message_text, 'html')
    else:
        msg1 = MIMEText(message_text, 'plain')
    message.attach(msg1)
    # message.attach(msg2)

    for file in attachments:
        content_type, encoding = mimetypes.guess_type(file)

        if content_type is None or encoding is not None:
            content_type = 'application/octet-stream'
        main_type, sub_type = content_type.split('/', 1)
        if main_type == 'text':
            fp = open(file, 'rb')
            msg = MIMEText(fp.read(), _subtype=sub_type)
            fp.close()
        elif main_type == 'image':
            fp = open(file, 'rb')
            msg = MIMEImage(fp.read(), _subtype=sub_type)
            fp.close()
        elif main_type == 'audio':
            fp = open(file, 'rb')
            msg = MIMEAudio(fp.read(), _subtype=sub_type)
            fp.close()
        elif main_type == 'application':
            fp = open(file, 'rb')
            msg = MIMEApplication(fp.read(), _subtype=sub_type)
            fp.close()
        else:
            fp = open(file, 'rb')
            msg = MIMEBase(main_type, sub_type)
            msg.set_payload(fp.read())
            fp.close()
        filename = os.path.basename(file)
        msg.add_header('Content-Disposition', 'attachment', filename=filename)
        message.attach(msg)
    return {'raw': base64.urlsafe_b64encode(message.as_string().encode()).decode()}


def create_draft(service, user_id, message_body):
    """Create and insert a draft email. Print the returned draft's message and id.

    Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message_body: The body of the email message, including headers.

    Returns:
    Draft object, including draft id and message meta data.
    """
    message = {'message': message_body}
    draft = service.users().drafts().create(userId=user_id, body=message).execute()

    print('Draft id: %s\nDraft message: %s' % (draft['id'], draft['message']))

    return draft


def build_service(credentials):
    """Build a Gmail service object.
    Args:
        credentials: OAuth 2.0 credentials.
    Returns:
        Gmail service object.
    """
    # credentials = get_credentials()
    http = httplib2.Http()
    http = credentials.authorize(http)
    return apiclient.discovery.build('gmail', 'v1', http=http)


if __name__ == '__main__':
    print('Gmail Access')
