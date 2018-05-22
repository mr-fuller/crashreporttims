import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

def composeAndSendEmail(files):

    FROM = 'fuller@tmacog.org'
    TO = 'mrfuller460@gmail.com' # vondeylen@tmacog.org, householder@tmacog.org
    SUBJECT = 'Crash Report Locations'
    TEXT = 'See attached Excel files'
    msg = MIMEMultipart()
    msg['From'] = FROM
    msg['To'] = TO
    msg['Subject'] = SUBJECT
    msg.attach(MIMEText(TEXT))
    #file = os.path.join(GDBspot, TimeDateStr + ".xlsx")
    # why does this work?
    for file in files:
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(file, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename= '+file)
        msg.attach(part)

    #msg.attach(file)

    # Connect to email server and send email
    with smtplib.SMTP('smtp.office365.com', 587) as s:
        s.starttls()
        s.login('fuller@tmacog.org','L7w##nie')
        s.send_message(msg)

if __name__ == "__main__":
    composeAndSendEmail()