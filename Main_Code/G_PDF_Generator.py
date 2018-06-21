"""This Moddule Searches People Table and Hesburgh Lecture Events Table in Access, as well as
   the clubs data google sheet to create and save Confirmation and Invoice PDFs
   The Module is called through the main program and is not directly executed"""

from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import datetime
import textwrap


def confirmation_pdf(clubname, size, coor_name, coor_email, ffirst, fmiddle, flast, faddress1, faddress2, fcity,
                     fstate, fzip, fphone, femail, date2, date3, title, fprefix):
    # Parameters are the necessary information from Access to create Confirmation
    today = datetime.datetime.today().strftime('%B %d, %Y')
    club = '{0} [{1}]'.format(clubname, size)

    # Accounting for unfilled parameters in hesb_lect_events and people tables
    if fcity is not None and fstate is not None and fzip is not None:
        fcity_state_zip = '{0}, {1} {2}'.format(fcity, fstate, fzip)
    elif fcity is not None and fstate is not None and fzip is None:
        fcity_state_zip = '{0}, {1}'.format(fcity, fstate)
    else:
        fcity_state_zip = ''

    if fprefix is not None and fmiddle is not None:
        faculty = '{0} {1} {2} {3}'.format(fprefix, ffirst, fmiddle, flast)
    elif fmiddle is not None:
        faculty = '{0} {1} {2}'.format(ffirst, fmiddle, flast)
    elif fprefix is not None:
        faculty = '{0} {1} {2}'.format(fprefix, ffirst, flast)
    else:
        faculty = '{0} {1}'.format(ffirst, flast)

    if faddress1 is not None:
        faddress1 = faddress1
    else:
        faddress1 = ''

    if faddress2 is not None:
        faddress2 = faddress2
    else:
        faddress2 = ''

    if fphone is not None:
        fphone = fphone
    else:
        fphone = ''

    if femail is not None:
        femail = femail
    else:
        femail = ''

    packet = io.BytesIO()

    # create a new PDF with Reportlab
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFont('Times-Bold', 8)
    can.setFillColorRGB(.009, .078, .263)
    can.drawString(158, 642, club)
    can.drawString(80, 627, coor_name)
    can.drawString(80, 542, faculty)
    can.drawString(315, 692.5, today)
    can.setFont('Times-Roman', 8)
    can.setFillColorRGB(.009, .078, .263)
    can.drawString(383, 642, str(date3))
    # Title has to be textwrapped if too long
    y = 627
    for line in textwrap.wrap(title, 49):
        can.drawString(353, y, line)
        y -= 10
    # Formatting for Faculty information
    can.drawString(80, 615, coor_email)
    can.drawString(80, 530, faddress1)
    can.drawString(80, 518, faddress2)
    can.drawString(80, 506, fcity_state_zip)
    can.drawString(80, 494, fphone)
    can.drawString(80, 482, femail)
    can.save()

    # move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    # read your existing PDF
    existing_pdf = PdfFileReader(open(r'G:\Team Drives\Alumni All Staff\Hesburgh Lecture Series\Confirmation Packet\Confirmations\Confirmation Template.pdf', "rb"))
    output = PdfFileWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    # finally, write "output" to a real file
    if '/' in clubname:
        clubname = clubname.replace('/', '_')
    else:
        clubname = clubname
    outputStream = open(r'G:\Team Drives\Alumni All Staff\Hesburgh Lecture Series\Confirmation Packet\Confirmations\rpt3HL_Confirm_' + clubname + '_' + date2 + '.pdf', "wb")
    output.write(outputStream)
    outputStream.close()
    output_loc = r'G:\Team Drives\Alumni All Staff\Hesburgh Lecture Series\Confirmation Packet\Confirmations\rpt3HL_Confirm_' + clubname + '_' + date2 + '.pdf'
    # output_loc is the location of the newly created confirmation. It is required for creation for corrdinator and
    # faculty confirmation emails
    return output_loc


def invoice_pdf(clubname, region, size, fprefix, ffirst, fmiddle, flast, date2,
                date3, title, payment, coor_name, due_date):
    # Parameters are the necessary information from Access to create Invoice
    if region is not None and size is not None:
        club = '{0} [{1}, {2}]'.format(clubname, region, size)
    elif region is not None:
        club = '{0} [{1}, ]'.format(clubname, region)
    elif size is not None:
        club = '{0} [ , {1}]'.format(clubname, size)
    else:
        club = '{0} [ , ]'.format(clubname)

    if fprefix is not None and fmiddle is not None:
        faculty = '{0} {1} {2} {3}'.format(fprefix, ffirst, fmiddle, flast)
    elif fmiddle is not None:
        faculty = '{0} {1} {2}'.format(ffirst, fmiddle, flast)
    elif fprefix is not None:
        faculty = '{0} {1} {2}'.format(fprefix, ffirst, flast)
    else:
        faculty = '{0} {1}'.format(ffirst, flast)

    if payment is not None:
        payment = payment
    else:
        payment = 0

    today = datetime.datetime.today().strftime('%B %d, %Y')
    packet = io.BytesIO()

    # create a new PDF with Reportlab
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFont('Times-Bold', 12.04)
    can.setFillColorRGB(.763, .59, 0)
    can.drawString(82, 553, club)
    can.setFont('Times-Bold', 10.02)
    can.setFillColorRGB(.009, .078, .263)
    can.drawString(383, 532, coor_name)
    can.drawString(143, 532, faculty)
    can.setFont('Times-Bold', 11.03)
    can.drawString(484, 458.5, '$' + str(payment) + '.00')
    can.setFont('Times-Bold', 9.01)
    can.drawString(315, 692.5, today)
    can.setFont('Times-Roman', 10.02)
    can.setFillColorRGB(.009, .078, .263)
    can.drawString(143, 517.25, date3)
    # Title has to be textwrapped if too long
    y = 502
    for line in textwrap.wrap(title, 35):
        can.drawString(143, y, line)
        y -= 15
    can.drawString(383, 517, '$' + str(payment) + '.00')
    can.drawString(383, 502, due_date)
    can.save()

    # move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    # read your existing PDF
    existing_pdf = PdfFileReader(open(r'G:\Team Drives\Alumni All Staff\Hesburgh Lecture Series\Confirmation Packet\Invoices\Invoice Template.pdf', "rb"))
    output = PdfFileWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    # finally, write "output" to a real file
    if '/' in clubname:
        clubname = clubname.replace('/', '_')
    else:
        clubname = clubname
    outputStream = open(r'G:\Team Drives\Alumni All Staff\Hesburgh Lecture Series\Confirmation Packet\Invoices\rpt3HL_Invoice_' + clubname + '_' + date2 + '.pdf', "wb")
    output.write(outputStream)
    outputStream.close()
    output_loc = r'G:\Team Drives\Alumni All Staff\Hesburgh Lecture Series\Confirmation Packet\Invoices\rpt3HL_Invoice_' + clubname + '_' + date2 + '.pdf'
    # output_loc is the location of the newly created invoice. It is required for creation for corrdinator
    # confirmation emails
    return output_loc


if __name__ == '__main__':
    print('PDF Generator')
