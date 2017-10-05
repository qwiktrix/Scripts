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

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from email.mime.text import MIMEText

try: 
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None
    
SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Tronlab Registration Application'
COORDINATOR_EMAIL = 'tronlab.coordinator@gmail.com'
PASSWORD = 'Mechatronics2010'

"""
Setting Variables
"""
SPREADSHEET_ID = '1xhiyMpNMcaolPF6UdTY4m5krlJpiAWHK3gNZwnURLHQ'
DEST_EMAIL = 'tronlab.coordinator@gmail.com'             #coordinator email - sends email to this acc
COLS_WITHOUT_CHK = 18                                           #total num of col w/o added check
FLAG_INDEX = 18                                                 #check col index
NAMES_INDEX = 13                                                #name col index
STUDENT_NUM_INDEX = 15                                          #student num index
DEPARTMENT_INDEX = 14                                           #program and year index
SCORE_INDEX = 16                                                #quiz score index
FORM_ID = 'Form Responses 1'                                    #form id /see below spreadsheet pg, the tab/
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
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:
            credentials = tools.run_flow(flow, store)
        print('Storing credentials to' + credential_path)
    return credentials

def isValid(chunk):
    if (chunk[SCORE_INDEX] == '10 / 10'):
        return True
    else:
        return False

def chk_extraction():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryURL = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets','v4', http=http, discoveryServiceUrl=discoveryURL)
    rangeName= FORM_ID+'!A2:S'
    result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID,range=rangeName).execute()
    values = result.get('values', [])
    
    emailChunk = []
    emailList = []
    failedList = []
        
    if not values:
        print('No data found.')
    else:       
        for i in range(len(values)):
            if (len(values[i]) == COLS_WITHOUT_CHK):
                studentRecord = values[i]
                
                if (isValid(studentRecord)):
                    studentChunk = [studentRecord[NAMES_INDEX], studentRecord[STUDENT_NUM_INDEX], studentRecord[DEPARTMENT_INDEX]] 
                    emailChunk.append(studentChunk)
                    emailList.append(i+2)
                else:
                    failedList.append(i+2)

        extract = [emailChunk, emailList, failedList]
        return extract
                           
    
def isEmpty(chunk):
    if (len(chunk[0]) == 0):
        return True
    else:
        return False 

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

def email(students):
    emailString = '<html><pre><font size= 3>Hello!\r\nThese are the students that need to be updated to have access to the Tronlab.\r\n\r\n' + students + '\r\nThank you.\r\nRegards,\r\nJessa</font></pre></html>'
    
    msg = MIMEText(emailString,'html')
    msg['Subject'] = 'Tronlab Access Confirmation'
    msg['To'] = COORDINATOR_EMAIL
    
    smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpObj.ehlo()
    smtpObj.starttls()
    smtpObj.login(COORDINATOR_EMAIL, PASSWORD)
    smtpObj.sendmail(COORDINATOR_EMAIL, DEST_EMAIL , msg.as_string())
    smtpObj.quit()
    
def updateSheet(target,flag):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryURL = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets','v4', http=http, discoveryServiceUrl=discoveryURL)
    rangeName = target
    valueBody = { "values": [[flag]]}
    update = service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID,range=rangeName, valueInputOption='RAW', body=valueBody).execute()
    
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
    