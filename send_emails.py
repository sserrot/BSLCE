import find_non_attendants
import smtplib  # import smtplib to send email

nonAttendantStudents = find_non_attendants.main()

# create a text/plain message


def send_emails(sender_address, password, subject, text_body):

    sender_address = sender_address
    password = password
    subject = subject
    text_body = text_body

    sender_address = '@gmail.com'
    password = 'password'
    subject = 'hello'
    text_body = 'testing'

    smtp_obj = smtplib.SMTP('smtp.gmail.com', 587)
    smtp_obj.ehlo()
    smtp_obj.starttls()
    smtp_obj.login(sender_address, password)

# calling str on text_body creates problems with unicode ... TODO - fix this
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
    send_emails(sender_address=sender_address, password=password, subject=subject, text_body=text_body)

if __name__ == '__main__':
    main(file='foo.xlsx', sender_address='no@gmail.com', password='pass', subject='subj', text_body='text_body')
