# Importing Libraries
import extract_msg
from openpyxl import Workbook, load_workbook
import openpyxl 
import os
import glob

# Taking emails from folders
emails = glob.glob("*.msg")
# Saving to excel
wb = openpyxl.Workbook()
ws1 = wb.active
# Inserting headers
headers = ["Email Sender", "Email Receiver", "Email Sent On", "Email Subject"]  
ws1.append(headers)

for email in emails:
    # File
    file = email
    # open message
    msg = extract_msg.Message(file)
    # print sender name
    'Sender: {}'.format(msg.sender)
    # print receivers name
    'Receiver: {}'.format(msg.to)
    # print date
    'Sent On: {}'.format(msg.date)
    # print subject
    'Subject: {}'.format(msg.subject)
    # # print body
    # print('Body: {}'.format(msg.body))
    
    data = [msg.sender, msg.to, msg.date, msg.subject]
    print(data)
    ws1.append(data)
    wb.save("Emails_Data.xlsx")