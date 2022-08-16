import smtplib
from config import *
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

username = EMAIL_USER
password = EMAIL_PASSWORD
mail_from = EMAIL_USER
mail_to = RECIPIENT
mail_subject = 'Database Successfully Updated - ' + TODAY 
mail_body = 'The database was successfully updated at ' + TODAY

try:
	mimemsg = MIMEMultipart()
	mimemsg['From']=mail_from
	mimemsg['To']=mail_to
	mimemsg['Subject']=mail_subject
	mimemsg.attach(MIMEText(mail_body, 'plain'))
	connection = smtplib.SMTP(host='smtp.office365.com', port=587)
	connection.starttls()
	connection.login(username,password)
	connection.send_message(mimemsg)
	connection.quit()
except Exception as error:
    print(error)