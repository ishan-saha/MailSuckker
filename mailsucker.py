from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from mimetypes import MimeTypes
import smtplib, ssl, time, hashlib, openpyxl
from bs4 import BeautifulSoup

# settings SMTP configuration
port = 465
username = '<email ID>' # the email provider can be any service like godaddy or porotonmail 
password = '<mail password>'
server = 'smtpout.secureserver.net' # smpt server 

# loading the HTML email teamplate the identifier id {{ID}}
file = open('email.html','r') # need to have this file as the email content 
soup = BeautifulSoup(file,features="html.parser")
file.close()

email_body = soup.prettify()

workbook = openpyxl.load_workbook('emails.xlsx') # this file contains the targets emails 
sheet = workbook.active
for row in range(0,sheet.max_row):
    for col in sheet.iter_cols(1,sheet.max_column):
        email=col[row].value
        if email== None:
            exit()
        # replacing the Identifier with a hash for proper identification
        body = email_body.replace('{{ID}}',hashlib.md5(email.encode()).hexdigest(),2) 
        print('mail sent to :',col[row].value,'identifier is:',hashlib.md5(email.encode()).hexdigest())
        # creating the message body
        message = MIMEMultipart('alternative')
        message['Subject'] = "Diwali Bonus FY 2022 | Triveni Group"
        message['From']=username
        message['To']=email
        bodytext = MIMEText(body,'html')
        message.attach(bodytext)
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(server,port, context=context) as Sucker:
                Sucker.login(username,password)
                Sucker.sendmail(username,email,message.as_string())
        except Exception as E:
            print(E)
        time.sleep(3)
