""" ------------------------------------------------------
Tronlab Registration Spreadsheet to Email Script
Written by Jessa Mae Alcantara
Last Updated: 10/04/2017

Code excerpts taken from Google Spreadsheets quickstart.py Eample from here : https://developers.google.com/sheets/api/quickstart/python
-----------------------------------------------------------
HOW TO USE
-----------------------------------------------------------
1) Ensure that 'client_secret.json' file is in same directory as this script
2) Ensure that all predefined settings follow your indicated spreadsheet 
   as seen below in "Setting Variables"
3) Run in command line as "python tronlab.py" and you're good to go!
-----------------------------------------------------------
"""
import httplib2
import os
import smtplib

import discovery
import client
import tools
from oauth2client.file import Storage
from email.mime.text import MIMEText

try: 
    import argparse
    flags = argparse.ArgumentParser(parents=[oauth2client.tools.argparser]).parse_args()
except ImportError:
    flags = None
    
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Tronlab Registration Application'
COORDINATOR_EMAIL = 'coordinator_email'
PASSWORD = 'password'

"""
Setting Variables
"""
SPREADSHEET_ID = 'SpreadsheetID'
DEST_EMAIL = 'destination_email'                                #coordinator email - sends email to this acc
AUTHOR = 'Tronlab Coordinator Name'                                                #coordinator name           
EMAIL_SUBJECT = 'Tronlab Access Confirmation'                   #email subject 
COLS_WITHOUT_CHK = 18                                           #total num of col w/o added check
FLAG_INDEX = 18                                                 #check col index
NAMES_INDEX = 13                                                #name col index
STUDENT_NUM_INDEX = 15                                          #student num index
DEPARTMENT_INDEX = 14                                           #program and year index
SCORE_INDEX = 16                                                #quiz score index
FORM_ID = 'SheetID'                                             #form id /see below spreadsheet pg, the tab/
CHK_COL_ID = 'S'                                                #form value for check col
PASS_VALUE = 'x'                                                #pass/fail value to insert
FAIL_VALUE = 'o'


def get_credentials():
    home = os.path.expanduser('~')
    credential_dir = os.path.join(home, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'sheets.googleapis.com-tronlab-registration.json')
    store = Storage(credential_path)
    credentials = store.get()
    
    if not credentials or credentials.invalid:
        flow = oauth2client.client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = oauth2client.tools.run_flow(flow, store, flags)
        else:
            credentials = oauth2client.tools.run_flow(flow, store)
        print('Storing credentials to' + credential_path)
    return credentials


# Checking the entry score is perfect
# NOTE: student may only have access IF a perfect score is attained
def isValid(chunk):
    if (chunk[SCORE_INDEX] == '10 / 10'):
        return True
    else:
        return False

# Checks document for all spreadsheet records that have not been accounted for, either that have not been emailed
# or marked as fail on spreadsheet document. If PASS {Name, Student Num, Department} extracted and added to emailChunk 
# and spreadsheet record index added to emailList, if FAIL index added to failedList
def chk_extraction():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryURL = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = apiclient.discovery.build('sheets','v4', http=http, discoveryServiceUrl=discoveryURL)
    rangeName= FORM_ID+'!A2:S'
    result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID,range=rangeName).execute()
    values = result.get('values', [])
    
    emailChunk, emailList, failedList = [], [], []
        
    if not values:
        print('No data found.')
    else:       
        # Checks and extracts for records that have not been checked no "Emailed" spreadsheet value, dictated by COLS_WITHOUT_CHK definition
        for i in range(len(values)):
            if (len(values[i]) == COLS_WITHOUT_CHK):
                studentRecord = values[i]
                
                # Check for score to be perfect
                if (isValid(studentRecord)):
                    studentChunk = [studentRecord[NAMES_INDEX], studentRecord[STUDENT_NUM_INDEX], studentRecord[DEPARTMENT_INDEX]] 
                    emailChunk.append(studentChunk)
                    emailList.append(i+2)
                else:
                    failedList.append(i+2)

        extract = [emailChunk, emailList, failedList]
        return extract
                

# Checks whether extration if all students have not been accounted for
# if so not necessary to email or if any failed scores students to update the sheet
def isEmpty(chunk):
    if (len(chunk[0]) == 0):
        return True
    else:
        return False 

# Formatting the data from spreadsheet {Name, Student Number, Department/Year} into columns for email chunk
# with proper padding
def formatChunk(Extraction):
        MaxNameLen = 0
        MaxStNumLen = 15
        studentData = ''
        
        for record in Extraction:
            if (len(record[0]) > MaxNameLen): 
                MaxNameLen = len(record[0])
            
        for record in Extraction:
            studentData = studentData + record[0] + (MaxNameLen - len(record[0])) * ' '  + '\t\t' + record[1] + (MaxStNumLen - len(record[1])) * ' ' + '\t\t\t' + record[2] + '\r\n'
        chunk = 'Name:' + (MaxNameLen - 3) * ' ' + '\t\tStudent Number:\t\t\tDepartment:\n' + studentData 
        return chunk

# Using SMTP email string with formatted student information as provided by formatChunk
#
# INPUTS    {  
#               students    =   formatted text of student data {Name, Student Number, Department/Year}
#           }
def email(students):
    emailString = '<html><pre><font size= 3>Hello!\r\nThese are the students that need to be updated to have access to the Tronlab.\r\n\r\n' + students + '\r\nThank you.\r\nRegards,\r\n' + AUTHOR + '</font></pre></html>'
    
    msg = MIMEText(emailString,'html')
    msg['Subject'] = EMAIL_SUBJECT
    msg['To'] = COORDINATOR_EMAIL
    
    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpObj.ehlo()
    smtpObj.starttls()
    smtpObj.login(COORDINATOR_EMAIL, PASSWORD)
    smtpObj.sendmail(COORDINATOR_EMAIL, DEST_EMAIL , msg.as_string())
    smtpObj.quit()

# Updates spreadsheet with marker in column referenced by "CHK_COL_ID" 
# using Google Spreadsheets API (note single cell write)
#
# INPUTS   {    
#               target  = spreadsheet cell to update 
#               flag    = character to fill in spreadsheet {PASS_VALUE, FAIL_VALUE}, as defined above          
#          }
def updateSheet(target,flag):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryURL = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = apiclient.discovery.build('sheets','v4', http=http, discoveryServiceUrl=discoveryURL)
    rangeName = target
    valueBody = { "values": [[flag]]}
    update = service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID,range=rangeName, valueInputOption='RAW', body=valueBody).execute()
    
# Iterates through indexList to update all necessary cells with a emailed flag value
#
# INPUTS   {    
#               indexList  = list of indexes of record that have to be updated b/w {emailList, failList} see chk_extraction
#               flag    = character to fill in spreadsheet {PASS_VALUE, FAIL_VALUE}, as defined above          
#          }
def writeTo(indexList,flag):
        for i in range(len(indexList)):
            upd_cell = FORM_ID + '!' + CHK_COL_ID + str(indexList[i])
            updateSheet(upd_cell,flag)
    
def main():
    print('Extracting unaccounted students ...')
    extraction = chk_extraction()
    
    print('Accounting for failed submissions ...')
    if (len(extraction[2]) != 0):
        writeTo(extraction[2], FAIL_VALUE)

    print('Checking if there are students to update ...')
    if (isEmpty(extraction)):
        print('All students have been accounted for. No need to update.')
    else:    
        email(formatChunk(extraction[0]))
        
        print('Updating spreadsheet ...')
        writeTo(extraction[1],PASS_VALUE)
            
main()
    