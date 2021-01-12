# libraries to be imported
import smtplib
import os
import glob
import xlwings as xw
import pandas as pd
import numpy as np
import io
import time


# first get the PDFs ready and named
# Next, make the email
# Then add the PDFs as attachments
# Add the tables into the email AS TEXT so people can copy!
# Send!  Or somehow keep it as a draft for me to review later?

# Get the name of the file. As there should only be *ONE* Excel file, make a list of one
path = r"C:\Users\haley\Python - Haley projects\RFI_Submittal_Email"
extensionXLSM = 'xlsm'
os.chdir(path)
resultXLSM = glob.glob('*.{}'.format(extensionXLSM))

# Run the marcos to make the PDFs and HTML files
wb = xw.Book(resultXLSM[0])  # remember, you made a list of 1, so you need the 0th item
app = xw.apps.active
macroRemoveOldFiles = wb.macro('removeOldStuff')
macroRFI = wb.macro('printPDF_HTML_RFIs')
macroSubmittal = wb.macro('printPDF_HTML_Submittals')
macroRemoveOldFiles()
macroRFI()
macroSubmittal()

# Quit Excel without saving (since we didn't do anything worth saving and just in case a macro goes wrong)
#wb.close() <--this closes the workbook but not the Excel app
app.quit()

# Finally, we need to know the names of the html files we just created, so lets make a list!
extensionHTML = 'html'
os.chdir(path + "\HTMLtables")  # go into the right folder
resultHTML = glob.glob('*.{}'.format(extensionHTML))

# Get the email stuff ready
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders

fromaddr = "TO@gmail.com"
toaddr = "FROM@gmail.com"

# instance of MIMEMultipart
msg = MIMEMultipart()

# storing the senders email address
msg['From'] = fromaddr

# storing the receivers email address
msg['To'] = toaddr

# storing the subject
msg['Subject'] = "Test email HTML"

# string to store the body of the mail
#body = "hello yes this is dog"
#msg.attach(MIMEText(body, 'plain'))


# Build the HTML body

os.chdir(path) #get the email base text from the main folder
part1Text = open("Email_text.html")

os.chdir(path + "\HTMLtables") #get the newly created HTML files from the HTML Holding folder

# gotta make sure the HTML files exist b/c this code runs faster than Excel's macros
attachmentsHTML = list(resultHTML)

for filename in attachmentsHTML:
    if len(attachmentsHTML) < 2:
        () #do nothing until it equals 2

part2Table1 = open(resultHTML[0])
part3Table2 = open(resultHTML[1])

# build the HTML message
htmlTablePretty = '''
<html>
<head>
<style></style>
</head>
<body>
<table>
<tr id="text">
'''
htmlTablePretty += part1Text.read()
htmlTablePretty += '''
</tr>
<tr id="RFI Table">
<td width=100% border = ".5" float:left style="width:100%!important; float:left">
'''
htmlTablePretty += part2Table1.read()
htmlTablePretty += '''
<p></p>
</td>
</tr>
<tr id="Submittal Table">
<td width=100% border = ".5" float:left style="width:100%!important; float: left">
'''
htmlTablePretty += part3Table2.read()
htmlTablePretty += '''
</td>
</tr>
</table>
</body>
</html>
'''

# os.chdir(path) #gotta make sure we're in the right directory
# fileMid=open('newEmail_Text.html','w')
# fileMid.write(htmlTablePretty)
# fileMid.close()

# put the pretty HTML into the email message
msg.attach(MIMEText(htmlTablePretty, 'html'))


# Add attachments

# Make a list of the PDFs so we can get a handle to them
extensionPDF = 'pdf'
os.chdir(path + "\PDFs")  # go into the PDF Holding folder
resultPDF = glob.glob('*.{}'.format(extensionPDF))
# force resultPDF to be a list that the loop can read
attachmentsPDF = list(resultPDF)

# start attachment for loop

# gotta make sure the PDF files exist b/c this code runs faster than Excel's macros
for filename in attachmentsPDF:
    if len(attachmentsPDF) < 2:
        () #do nothing until it equals 2
    else:
        f = filename
        part = MIMEBase('application', "octet-stream")
        part.set_payload( open(f,"rb").read() )
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachments; filename="%s"' % os.path.basename(f))
        msg.attach(part)


# creates SMTP session
s = smtplib.SMTP('smtp.gmail.com', 587)

# start TLS for security
s.starttls()

# Authentication
s.login(fromaddr, "password")

# Converts the Multipart msg into a string
text = msg.as_string()

# sending the mail
s.sendmail(fromaddr, toaddr, text)

# terminating the session
s.quit()
