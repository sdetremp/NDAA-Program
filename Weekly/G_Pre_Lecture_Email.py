"""
   Hesburgh Lecture Series Processing Program -- Creator: Sam DeTrempe, Class of 2018

   This program generates emails to be sent out prior to Hesburgh Lecture events.
   This Program Uses 2 accompanying HLS specific modules and is a direct executable
"""

import pypyodbc
import G_Gmail_Access, Google_Spreadsheet
import datetime
import time

CON = pypyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb)};UID=admin;UserCommitSync=Yes;Threads=3;'
                       r'SafeTransactions=0;PageTimeout=5;MaxScanRows=8;MaxBufferSize=2048;FIL={MS Access};'
                       r'DriverId=25;DefaultDir=C:/Users/sdetremp/Python_MS_Access/Database;'
                       r'DBQ=P:\Databases\Academic Databases\AcademicPrograms_apps_Robby.mdb;')

CURSOR = CON.cursor()


def search_lectures_and_draft_email():
    # Searches events table in Access for lectures that are coming up
    CURSOR.execute('SELECT * FROM hesb_lect_events ORDER BY [HL_Events_ID] DESC')
    # Ask user how many days in advance to search for upcoming lectures
    user_input = input('How many days in advance would you like to create Pre-Lecture notification emails: ')
    today = datetime.datetime.now()
    today2 = datetime.datetime.today().strftime('%m/%d/%Y')
    week_ahead = today + datetime.timedelta(days=int(user_input))

    lectures = CURSOR.fetchall()

    for row in lectures:
        try:
            lect_date = row['date of lecture']
            if lect_date <= week_ahead and lect_date > today and \
                    row['co-sponsor contact 2 faxphone'] is None and \
                    row['status'] is None:
                rowid = str(row['hl_events_id'])
                clubname = row['clubsregionsid']
                header_date = lect_date.strftime('%m/%d/%Y')
                body_date = lect_date.strftime('%B %d')
                lecturer = row['hl_resource']
                coordinator = row['club contact']

                # Access people table and pull necessary info
                CURSOR.execute('SELECT * FROM people')
                prow = CURSOR.fetchall()
                for i in prow:
                    # Lecturer Info
                    if i['id'] == coordinator:
                        lfirst = i['firstname']
                        llast = i['lastname']
                        cemail = i['email1']
                    # Faculty Info
                    elif i['id'] == lecturer:
                        ffirst = i['firstname']
                        flast = i['lastname']
                        fprefix = i['prefix']
                        femail = i['email1']

                # Access Clubs Regions Table and pull club number
                CURSOR.execute('SELECT * FROM tblClubsRegions')
                crow = CURSOR.fetchall()
                for i in crow:
                    if i['clubsregionsid'] == clubname:
                        clubnumber = i['clubnumber']
                        break
                    else:
                        pass
                # Club number from Access is used to pull info from Clubs Data Google Sheet
                clubname, region, size = Google_Spreadsheet.get_clubinfo(clubnumber)
                coor_name = (lfirst + ' ' + llast)

                draft_email(clubname, header_date, body_date, coor_name, cemail, fprefix, ffirst, flast, femail)
                CURSOR.execute("UPDATE hesb_lect_events SET "
                               "[co-sponsor contact 2 faxphone]='Sent - " + str(today2) +
                               "' WHERE [HL_Events_ID]=" + rowid + "")
                CON.commit()
        except Exception as e:
            print(e)
    print('All Pre-Lecture Emails Drafted')
    print('Terminating Process')
    time.sleep(4)


def draft_email(clubname, header_date, body_date, coor_name, cemail, fprefix, ffirst, flast, femail):
    # Draft Pre Lecture Email
    # Body of email is pulled from text file in G drive. Email is written using html
    if fprefix == 'Reverend':
        ftitle = 'Father'
    else:
        ftitle = 'Professor'

    subject = '2018 Upcoming Hesburgh Lecture for ND Club of ' + clubname + ', ' + header_date
    with open(r'G:\Team Drives\Alumni All Staff\Hesburgh Lecture Series\_HLS Code - DO NOT DELETE OR MOVE\Pre-Lecture_Email_Content.txt', encoding='utf8') as f:
        message_text = f.read().replace('coor_name', coor_name).replace('ftitle', ftitle)\
                               .replace('flast', flast).replace('ffirst', ffirst)\
                               .replace('clubname', clubname).replace('body_date', body_date)

    files = []
    draft = G_Gmail_Access.create_message_with_attachment('alumaced@nd.edu', str(cemail) + ',' + str(femail), '',
                                                        subject, message_text, files, verbose=True)
    service = G_Gmail_Access.build_service(G_Gmail_Access.get_credentials())
    G_Gmail_Access.create_draft(service, 'me', draft)
    return


if __name__ == '__main__':
    search_lectures_and_draft_email()
