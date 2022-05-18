import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta
from tzlocal import get_localzone

local_tz = get_localzone()   # get local timezone
now = datetime.now(local_tz) # get timezone-aware datetime object
today = now.strftime('%Y-%m-%d %H:%M')

username = MMG_USER
password = MMG_PASSWORD
mail_from = MMG_USER
mail_to = RECIPIENT
mail_subject = 'Database Successfully Completed - ' + today 
mail_body = 'The database was successfully updated at ' + today

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