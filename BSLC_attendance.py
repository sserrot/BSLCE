import openpyxl
from openpyxl import load_workbook # import openpyxl for excel manipulation
import smtplib # import smtplib to send email


# TO DO - abstract column 4 part of code and (D) part
nonAttendantStudents = {}


def excel_parse(file):
    wb = load_workbook(file, use_iterators=True)
    sheet = wb.get_sheet_by_name(name='Sheet1')
    i = 1
    for row in sheet.iter_rows():
        if sheet.cell(row=i, column=1).value is not None:
            for cell in row:
                print(i)
                print(sheet.cell(row=i, column=1).value)
                if str(sheet.cell(row=i, column=4).value).upper() == 'MISS':  # check if person attended
                    full_name = sheet.cell(row=i, column=1).value  # store full name
                    print(full_name)
                    last = full_name.rsplit(' ')[0].strip(',')  # format - broken if MISS is empty
                    if len(full_name.rsplit(' ')) == 1:
                        continue
                    else:
                        first = full_name.rsplit(' ')[1]  # format, causes an error if not in LAST, FIRST format
                        email = (
                        last[:7] + '_' + first[:4] + '@bentley.edu')  # email format is 7last_2first@bentley.edu
                        name = first + ' ' + last
                        nonAttendantStudents[name] = email  # add to a dictionary
                        i += 1
                        continue
                else:
                    i += 1  # increment where you are checking
                    continue
        else:
            break
# create a text/plain message


def send_emails(sender_address, password, subject, text_body):

    sender_address = sender_address
    password = password
    subject = subject
    text_body = text_body

    smtp_obj = smtplib.SMTP('smtp.gmail.com', 587)
    smtp_obj.ehlo()
    smtp_obj.starttls()
    smtp_obj.login(sender_address, password)

    for name, email in nonAttendantStudents.items():
        body = ("Subject:" + str(subject) + "\n" + "Dear {}, \n\n" + str(text_body)).format(name)
        print('Sending email to %s...' % email)
        send_mail_status = smtp_obj.sendmail(sender_address, email, body)

        if send_mail_status != {}:
            print('There was a problem sending email to %s: %s' % (email, send_mail_status))
    all_body = "Subject:" + str(subject) + "\n" + "Dear BSLCE, \n\n"
    for name in nonAttendantStudents.items():
        all_body = all_body + ' ' + name[0] + ";"

    send_all_status = smtp_obj.sendmail(sender_address, sender_address, all_body)

    if send_all_status != {}:
        print('There was a problem sending email to %s: %s' % (sender_address,send_all_status))
    smtp_obj.quit()


def main(file, sender_address, password, subject, text_body):
    excel_parse(file=file)
    send_emails(sender_address=sender_address, password=password, subject=subject, text_body=text_body)

if __name__ == '__main__':
    main(file='foo.xlsx', sender_address='no@gmail.com', password='pass', subject='subj', text_body='text_body')
