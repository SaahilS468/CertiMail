import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
import os
from dotenv import load_dotenv
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
load_dotenv()
# wdFormatPDF = 17

# Configuration
template = DocxTemplate('Contribution.docx')
filename = 'Event head.csv'
output = '/Certificates'

sender_address = os.getenv("SENDER")
sender_pass = os.getenv("PASS")


data = pd.read_csv(filename)
names = data["Name"]
email = data["Email"]

def create_certification():
    for name in names:
        context = {
            'Name': name
        }
        template.render(context)
        template.save(f".{output}/{name}.docx")
        outputPath = f".{output}/{name}.docx"
        pdfpath = outputPath[:-4] + 'pdf'
        convert(outputPath, pdfpath)


    inp = int(input("Enter 1 to mail: "))
    if inp:
        send_email(sender_address, names, email)
    else:
        print("Only Certificates will be generated")


def send_email(sender_address, names, email):

    for i in range(len(names)):

        person_name = names[i]

        message = MIMEMultipart()
        message['From'] = (f"CSI"
                           f"<{sender_address}>")
        message['To'] = email[i]
        message['Subject'] = 'Certificate for Participating in Competition '

        mail_content = f"""
        Dear {person_name},
    
        Thank you for participating in the competition.
    
        Find attached your certificate.
    
        Wishing you the best for all future endeavors	
    
        Regards,    
        """

        message.attach(MIMEText(mail_content, "plain"))
        pdf_path = os.getcwd() + output + "/" + " ".join(str(names[i]).split(' ')) + '.pdf'
        with open(pdf_path, "rb") as f:
            certi = MIMEApplication(f.read(), _subtype="pdf")
        certi.add_header('Content-Disposition', 'attachment', filename=str(names[i]))
        message.attach(certi)

        text = message.as_string()

        # Log in to server using secure context and send email
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(sender_address, sender_pass)
                server.sendmail(sender_address, email[i], text)
            print(f'Mail Sent to {person_name}')
        except smtplib.SMTPAuthenticationError:
            print("The username and/or password you entered is incorrect")
        except:
            print(f'Failed to send mail to {person_name}')

# Run function
create_certification()
