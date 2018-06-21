"""This module is used to update and pull info from NDAA specific google sheets
   The Module is called through the main program and is not directly executed"""

import gspread
from oauth2client.service_account import ServiceAccountCredentials

SCOPE = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
CREDS = ServiceAccountCredentials.from_json_keyfile_name(
    'G:\Team Drives\Alumni All Staff\Hesburgh Lecture Series\_HLS Code - DO NOT DELETE OR MOVE\_credentials\spreadsheet_secret.json',
    SCOPE)
CLIENT = gspread.authorize(CREDS)


def update_hl_sheet(club, date, ffirst, flast, title, lect_title):
    # Parameters are the necessary info used to update MY ND Events Calendar Google Sheet with Hesburgh Lecture info
    sheet = CLIENT.open('MyND Events Calendar List').sheet1

    row = ['Notre Dame Club of ' + club + ' Hesburgh Lecture', date, '', '', '', 'Speaker: '
           + ffirst + ' ' + flast + ', ' + title + '    Lecture: ' + lect_title, 'Learn, Connect']
    index = 4
    sheet.insert_row(row, index)


def get_clubinfo(club_num):
    # Reads in club information from Clubs Database Google Sheet
    # club_num is the club number pulled from Access. It is then used to look up information for the
    # corresponding club
    sheet = CLIENT.open_by_url(
        'https://docs.google.com/spreadsheets/d/1oXC29hV6ggrHlhb-vnSY008cs4uWEKAagHaH4BuoJpE/edit?usp=sharing')
    club_sheet = sheet.worksheet('ClubData')
    row = club_sheet.find('CCG00' + club_num).row
    # Get Column numbers for club name, region, size
    club_col = club_sheet.find('Club Name').col
    reg_col = club_sheet.find('Region').col
    size_col = club_sheet.find('Size').col

    club_name = club_sheet.cell(row, club_col).value
    reg = club_sheet.cell(row, reg_col).value
    size = club_sheet.cell(row, size_col).value
    return club_name, reg, size


def get_coor_info(club_num):
    # Reads in HLS coordinator information from Clubs Database Google Sheet
    # club_num is the club number pulled from Access. It is then used to look up information for the
    # corresponding club
    sheet = CLIENT.open_by_url(
        'https://docs.google.com/spreadsheets/d/1oXC29hV6ggrHlhb-vnSY008cs4uWEKAagHaH4BuoJpE/edit?usp=sharing')
    coor_sheet = sheet.worksheet('ClubOfficers')
    row = coor_sheet.find('CCG00' + club_num).row

    name_col = coor_sheet.find('CAR-Hesburgh_Lecture').col
    name = coor_sheet.cell(row, name_col).value

    email_col = coor_sheet.find('CAR-Hesburgh_Lecture_Coordinator_Email').col
    email = coor_sheet.cell(row, email_col).value
    return name, email


def insert_coor(club_num, coor_name, cemail):
    # Replaces current HLS coordinator name and email with new coordinator name and email when prompted
    # club_num is the club number pulled from Access. It is then used to look up information for the
    # corresponding club
    # coor_name is the coordinator name in the HLS email request
    # cemail is the coordinator email in the HLS email request
    sheet = CLIENT.open_by_url(
        'https://docs.google.com/spreadsheets/d/1oXC29hV6ggrHlhb-vnSY008cs4uWEKAagHaH4BuoJpE/edit?usp=sharing')
    coor_sheet = sheet.worksheet('ClubOfficers')
    row = coor_sheet.find('CCG00' + club_num).row

    name_col = coor_sheet.find('CAR-Hesburgh_Lecture').col
    name = coor_sheet.cell(row, name_col).value

    email_col = coor_sheet.find('CAR-Hesburgh_Lecture_Coordinator_Email').col
    email = coor_sheet.cell(row, email_col).value

    coor_sheet.update_cell(row, name_col, coor_name)
    coor_sheet.update_cell(row, email_col, cemail)


if __name__ == '__main__':
    print('Google Sheets Module')
