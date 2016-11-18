import find_non_attendants
import smtplib  # import smtplib to send email

# create a text/plain message


def email(sender_address, password, subject, body_text, non_attendant_students):

    sender_address = sender_address
    password = password
    subject = subject
    body_text = body_text
    non_attendant_students = non_attendant_students


    smtp_obj = smtplib.SMTP('smtp.gmail.com', 587)
    smtp_obj.ehlo()
    smtp_obj.starttls()
    smtp_obj.login(sender_address, password)

# calling str on body_text creates problems with unicode ... TODO - fix this
    for name, email in non_attendant_students.items():
        body = ("Subject:" + str(subject) + "\n" + "Dear {}, \n\n" + str(body_text)).format(name)
        print('Sending email to %s...' % email)
        send_mail_status = smtp_obj.sendmail(sender_address, email, body)

        if send_mail_status != {}:
            print('There was a problem sending email to %s: %s' % (email, send_mail_status))
    all_body = "Subject:" + str(subject) + "\n" + "Dear self, \n\n"
    for name in non_attendant_students.items():
        all_body = all_body + ' ' + name[0] + ";"

    send_all_status = smtp_obj.sendmail(sender_address, sender_address, all_body)

    if send_all_status != {}:
        print('There was a problem sending email to %s: %s' % (sender_address,send_all_status))
    smtp_obj.quit()


def main(file, sender_address, password, subject, body_text):
    email(sender_address=sender_address, password=password, subject=subject, body_text=body_text)

if __name__ == '__main__':
    main(file='foo.xlsx', sender_address='no@gmail.com', password='pass', subject='subj', body_text='body_text', non_attendant_students='test')
