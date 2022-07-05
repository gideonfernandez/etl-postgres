import smtplib
from config import *
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


username = MMG_USER
password = MMG_PASSWORD
mail_from = MMG_USER
mail_to = 'analytics@montagemarketinggroup.com, ematel@montagemarketinggroup.com, mkreider@montagemarketinggroup.com, gfernandez@montagemarketinggroup.com'
mail_subject = 'Weekly Pyxis Issues Report - ' + TODAY_SLASH 
mail_body = 'Hi team, \n\nPlease find the weekly Pyxis Issues Report attached for activities starting June 1, 2022 and forward. \n\nRegards, \nGideon'

filepath = r'data/Pyxis Report Issues_' + TODAYSTR + '.xlsx'
filename = r'Pyxis Report Issues_' + TODAYSTR + '.xlsx'
	
try:
	mimemsg = MIMEMultipart()
	mimemsg['From']=mail_from
	mimemsg['To']=mail_to
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