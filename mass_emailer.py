import csv
import os
import smtplib
import ssl

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

#---[ Module Level Variables ]---#
strSMPTPSSLserver = "smtp.someServer.com"                                                                            #For SMTP with SSL server
intSMTPSSLport = 465                                                                                                 #For SMTP with SSL port number
strSMTPSSLSenderEmail = "someSender@someServer.com"                                                                  #For SMTP with SSL sender
strSMPTPserver = "smtp.someServer.com"                                                                               #For SMTP server
intSMTPport = 25                                                                                                     #For SMTP port number
strSMTPSenderEmail = "someSender@someServer.com"                                                                     #For SMTP sender
strPassword = input("Type your password and press Enter:")
strSubject = "An email with attachment from Python"
strPlainBody = "This is an email with attachment(s) sent from Python"                                                #Use this for a plain text message
strHTMLBody = "<html><head></head><body>This is an email with <b>attachment(s)</b> sent from Python</body></html>"   #Use this for an html message
strDirectoryPath = 'C:\\Users\\User1\\Downloads'                                                                     #Location of attachments
lstFileTypes = ['.pdf', '.doc', '.docx']                                                                             #Allowed file types as attachments

#---[ Function to Send Emails ]---#
def compose_send_email(strReceiverEmail, blnSSL=False):
    try:
        blnError = False
        #Prepare the header information of the email message
        message = MIMEMultipart()
        message["From"] = strSMTPSSLSenderEmail if blnSSL else strSMTPSenderEmail
        message["To"] = strReceiverEmail
        message["Subject"] = strSubject
        message["Bcc"] = strReceiverEmail
        # Add body to email
        # message.attach(MIMEText(strPlainBody, "plain"))    #Use this for a plain text message
        message.attach(MIMEText(strHTMLBody, "html"))        #Use this for an html message
        # Obtain all valid attachments for the email based on allowed File Types in lstFileTypes.
        lstAttachments = []
        for filename in os.listdir(strDirectoryPath):
            name, ext = os.path.splitext(filename)
            if ext in lstFileTypes:
                lstAttachments.append(filename)
        # Sort the attachments so that they are added to the email in sorted order.
        lstAttachments.sort()
        # Attach the attachments to the email.
        for attfilename in lstAttachments:
            strFilePath = os.path.join(strDirectoryPath, attfilename)
            attachment = MIMEApplication(open(strFilePath, "rb").read(), _subtype="txt")
            attachment.add_header("Content-Disposition", 'attachment', filename=attfilename)
            message.attach(attachment)
        # Log into server and send email
        if blnSSL:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(strSMPTPSSLserver, intSMTPSSLport, context=context) as server:
                server.login(strSMTPSSLSenderEmail, strPassword)
                server.sendmail(strSMTPSSLSenderEmail, strReceiverEmail, message.as_string())
        else:
            with smtplib.SMTP(strSMPTPserver) as server:
                server.set_debuglevel(False)
                #server.login(strSMTPSenderEmail, strPassword)
                server.sendmail(strSMTPSenderEmail, strReceiverEmail, message.as_string())
    except Exception as ex:
        blnError = True
        print("Error encountered while sending emails.", ex)
    finally:
        return not(blnError)

#---[ Main Function to Obtain Email Addresses for Recipients ]---#
if __name__ == "__main__":
    try:
        blnEmailSent = False
        with open('C:\\Users\\User1\\Downloads\\email_recipients.csv') as file:
            reader = csv.reader(file)
            next(reader)
            for firstName, lastName, emailAddress in reader:
                print(f"Sending email to: {firstName + ' ' + lastName}")
                blnEmailSent = compose_send_email(emailAddress, False)
                if not(blnEmailSent):
                    print(f"Unable to send email to: " + " " + emailAddress)
    except Exception as ex:
        print("Error encountered while obtaining recipient email addresses.", ex)
