"""
   Hesburgh Lecture Series Processing Program -- Creator: Sam DeTrempe, Class of 2018

   This program processes new Hesburgh Lecture requests by connecting to the hesb_lect_events
   and people tables in the Access Database. The program does the following:
   - Generates Confirmation and stores in
     G:\Team Drives\Alumni All Staff\Hesburgh Lecture Series\Confirmation Packet\Confirmations
   - Generates Invoice and stores in
     G:\Team Drives\Alumni All Staff\Hesburgh Lecture Series\Confirmation Packet\Invoices
   - Drafts Coordinator Email
   - Drafts Faculty Email
   - Updates Google Calendar for alumaced@nd.edu
   - Inserts entry for new lecture in the MyND Events Calendar List Google Spreadsheet
   This Module program uses 4 accompanying HLS related modules (called below) and is a direct executable
"""

import pypyodbc
import G_PDF_Generator, G_Google_Calendar, G_Gmail_Access, Google_Spreadsheet
import time

CON = pypyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb)};UID=admin;UserCommitSync=Yes;Threads=3;'
                       r'SafeTransactions=0;PageTimeout=5;MaxScanRows=8;MaxBufferSize=2048;FIL={MS Access};'
                       r'DriverId=25;DefaultDir=C:/Users/sdetremp/Python_MS_Access/Database;'
                       r'DBQ=P:\Databases\Academic Databases\AcademicPrograms_apps_Robby.mdb;')

CURSOR = CON.cursor()


def main():
    # Prompt user for number of HLS events to process and which parts to process
    print('Welcome to the Hesburgh Lecture Series Processing Program')
    user_input = input('Please select the number of events you would like to process: \n')
    which_parts = input('Which of the following would you like to generate: (Please separate numbers with a comma(,))'
                        '\n0: All'
                        '\n1: Confirmation'
                        '\n2: Invoice'
                        '\n3: Coordinator Email'
                        '\n4: Faculty Email'
                        '\n5: Calendar'
                        '\n6: Google Sheet Update\n')
    print('Processing events...')
    # For each lecture pull the necessary information from Access
    # Necessary information includes everything used in confirmation, invoice, calendar, google sheet, and email
    for lecture in range(int(user_input)):
        coordinator, faculty, clubname, date0, date1, date2, date3, title, payment, coor_name, cemail, ffirst, \
         fmiddle, flast, fprefix, faddress1, faddress2, fcity, fstate, fzip, fphone, femail, ftitle, region, size, \
         due_date, clubnumber = pull_info(lecture)
        print('Generating event for ND Club of ' + clubname)
        # Inform user of current HLS coordinator info in Clubs Data Google Sheet
        # Asks user if they would like to use a different coordinator
        coor_check = input('\nSuggested coordinator info:'
                           '\nName: ' + coor_name + ''
                           '\nEmail: ' + cemail + ''
                           '\nWould you like to proceed? [Y/N]\n')
        if coor_check == 'y' or coor_check == 'Y':
            pass
        else:
            # If user wishes to use a different coordinator, user is asked to enter coordinator
            # name (first and last) and email
            # User is then asked if this new coordinator info should replace the old corrdinator info
            # in the Clubs Data Google Sheet. If yes, a change to the sheet will be made
            coor_name = input('\nPlease enter new Coordinator Name (first and last with space inbetween): ')
            cemail = input('Please enter new Coordinator Email: ')
            update_google_sheet = input('\nWould you like to update the Clubs Database in'
                                        '\nGoogle Sheets with this information? [Y/N]\n')
            if update_google_sheet == 'y' or update_google_sheet == 'Y':
                Google_Spreadsheet.insert_coor(clubnumber, coor_name, cemail)
            else:
                pass
        if which_parts == '0':
            which_parts = ['3', '4', '5', '6']
        #     Different functions are called based on which processes are selected
        for i in list(which_parts):
            if i == '1':
                confirm_attachment = G_PDF_Generator.confirmation_pdf(clubname, size, coor_name, cemail, ffirst,
                                                                      fmiddle, flast, faddress1, faddress2, fcity,
                                                                      fstate, fzip, fphone, femail, date2, date3,
                                                                      title, fprefix)
            if i == '2':
                invoice_attachment = G_PDF_Generator.invoice_pdf(clubname, region, size, fprefix, ffirst, fmiddle,
                                                                 flast, date2, date3, title, payment, coor_name,
                                                                 due_date)
            if i == '3':
                confirm_attachment = G_PDF_Generator.confirmation_pdf(clubname, size, coor_name, cemail, ffirst,
                                                                      fmiddle, flast, faddress1, faddress2, fcity,
                                                                      fstate, fzip, fphone, femail, date2, date3,
                                                                      title, fprefix)
                invoice_attachment = G_PDF_Generator.invoice_pdf(clubname, region, size, fprefix, ffirst, fmiddle,
                                                                 flast, date2, date3, title, payment, coor_name,
                                                                 due_date)
                coor_email(clubname, date1, coor_name, cemail, confirm_attachment, invoice_attachment)
            if i == '4':
                confirm_attachment = G_PDF_Generator.confirmation_pdf(clubname, size, coor_name, cemail, ffirst,
                                                                      fmiddle, flast, faddress1, faddress2, fcity,
                                                                      fstate, fzip, fphone, femail, date2, date3,
                                                                      title, fprefix)
                faculty_email(clubname, date1, femail, flast, fprefix, confirm_attachment)
            if i == '5':
                G_Google_Calendar.create_event(ffirst, flast, clubname, date0)
            if i == '6':
                Google_Spreadsheet.update_hl_sheet(clubname, date1, ffirst, flast, ftitle, title)
        print('Event Generated')
    CON.close()
    print('All Events Generated\nTerminating Process')
    time.sleep(4)


def pull_info(lecture):
    # Access hes_lect_events table and pull necessary info
    CURSOR.execute('SELECT * FROM hesb_lect_events ORDER BY [HL_Events_ID] DESC')
    row = CURSOR.fetchall()
    row = row[lecture]
    print(row)
    coordinator = row['club contact']
    faculty = row['hl_resource']
    clubname = row['clubsregionsid']
    date = row['date of lecture']
    date0 = date.strftime('%Y-%m-%d')
    date1 = date.strftime('%m/%d/%Y')
    date2 = (date1[:2] + date1[3:5] + date1[8:10])
    date3 = date.strftime('%B %d, %Y')
    due_date = row['club payment due date']
    due_date = due_date.strftime('%B %d, %Y')
    title = row['title']
    payment = row['club payment amount']

    # Access people table and pull necessary info
    CURSOR.execute('SELECT * FROM people')
    prow = CURSOR.fetchall()
    for i in prow:
        # Faculty Info
        if i['id'] == faculty:
            ffirst = i['firstname']
            fmiddle = i['middleinitial']
            flast = i['lastname']
            fprefix = i['prefix']
            faddress1 = i['address1']
            faddress2 = i['address2']
            fcity = i['city']
            fstate = i['state']
            fzip = i['zip']
            fphone = i['busphone1']
            femail = i['email1']
            ftitle = i['title1']
        else:
            pass

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
    coor_name, cemail = Google_Spreadsheet.get_coor_info(clubnumber)

    return coordinator, faculty, clubname, date0, date1, date2, date3, title, payment, coor_name, cemail, ffirst, \
           fmiddle, flast, fprefix, faddress1, faddress2, fcity, fstate, fzip, fphone, femail, ftitle, region, size,\
           due_date, clubnumber


def coor_email(club, date, coor_name, cemail, confirm_attachment, invoice_attachment):
    # Create Coordinator Email with Attachments
    # Body of email is pulled from text file in G drive. Email is written using html
    subject = '2018 Confirmation for the ND Club of ' + club + ' Hesburgh Lecture, ' + date
    # Read email and replace key words with specific HLS information
    with open(r'G:\Team Drives\Alumni All Staff\Hesburgh Lecture Series\_HLS Code - DO NOT DELETE OR MOVE\Coordinator_Confirmation_Email_Content.txt', encoding='utf8') as f:
        message_text = f.read().replace('coor_name', coor_name)

    files = [confirm_attachment, invoice_attachment]
    draft = G_Gmail_Access.create_message_with_attachment('alumaced@nd.edu', cemail, '', subject, message_text, files,
                                                        verbose=True)
    service = G_Gmail_Access.build_service(G_Gmail_Access.get_credentials())
    G_Gmail_Access.create_draft(service, 'me', draft)
    return


def faculty_email(club, date, femail, lastname, prefix, confirm_attachment):
    # Create Faculty Email with Attachments
    # Body of email is pulled from text file in G drive. Email is written using html
    if prefix == 'Reverend':
        title = 'Father'
    else:
        title = 'Professor'
    subject = '2018 Faculty Confirmation for the ND Club of ' + club + ' Hesburgh Lecture, ' + date
    # Read email and replace key words with specific HLS information
    with open(r'G:\Team Drives\Alumni All Staff\Hesburgh Lecture Series\_HLS Code - DO NOT DELETE OR MOVE\Faculty_Confirmation_Email_Content.txt', encoding='utf8') as f:
        message_text = f.read().replace('f_title', title).replace('last_name', lastname)

    files = [confirm_attachment]
    draft = G_Gmail_Access.create_message_with_attachment('alumaced@nd.edu', femail, 'marciafewell@anthonytravel.com',
                                                        subject, message_text, files, verbose=True)
    service = G_Gmail_Access.build_service(G_Gmail_Access.get_credentials())
    G_Gmail_Access.create_draft(service, 'me', draft)
    return


if __name__ == '__main__':
    main()
