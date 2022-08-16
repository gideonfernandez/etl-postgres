import smtplib
from config import *
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

username = EMAIL_USER
password = EMAIL_PASSWORD
mail_from = EMAIL_USER
mail_to = RECIPIENT
mail_cc = 'recipient_3@mail.com'
mail_subject = 'Weekly Issues Report - ' + TODAY_SLASH 
mail_body = 'Hi all, \n\nPlease find the weekly Issues Report attached for activities starting from X date and forward. \n\nRegards, \nData Team'

filepath = ISSUES_FILEPATH
filename = ISSUES_FILENAME

try:
	mimemsg = MIMEMultipart()
	mimemsg['From']=mail_from
	mimemsg['To']=mail_to
	mimemsg['Cc']=mail_cc
	mimemsg['Subject']=mail_subject
	mimemsg.attach(MIMEText(mail_body, 'plain'))

	fp = open(filepath, 'rb')
	part = MIMEBase('application','vnd.ms-excel')
	part.set_payload(fp.read())
	fp.close()
	encoders.encode_base64(part)
	part.add_header('Content-Disposition', 'attachment', filename=filename)
	mimemsg.attach(part)

	connection = smtplib.SMTP(host='smtp.office365.com', port=587)
	connection.starttls()
	connection.login(username,password)
	connection.send_message(mimemsg)
	connection.quit()
except Exception as error:
    print(error)