from gmailAPI import main
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import xlrd

credentials = input('Enter the json file name of gmail api credentials. (Omit ".json"): ')
service = main(f'{credentials}.json')   #Enter the file name of gmail api credentials

def usingTXT(FileTXT,emailBody,emailSubject):
        collection_of_emails = open(f'{FileTXT}.txt')
        lines = collection_of_emails.readline()

        while(lines):
                lines = lines.replace("\xa0",'')
                splitlines = lines.split('\t')
                recieverEmail = splitlines[1]
                recieverName = splitlines[0]
                subjectEmail = f'Hello {recieverName} {emailSubject}'
                bodyEmail = emailBody
                mimeMessage = MIMEMultipart()
                mimeMessage['to'] = recieverEmail
                mimeMessage['subject'] = subjectEmail
                mimeMessage.attach(MIMEText(bodyEmail, 'plain'))
                raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()
                message = service.users().messages().send(userId='me', body={'raw': raw_string}).execute()
                print(f'Message sent successfully to {recieverName}\n{message}\n')

                lines = collection_of_emails.readline()
                
                collection_of_emails.close()

##############################################################################################################

def usingXLSX(FileXLSX,emailBody,emailSubject):
        book = xlrd.open_workbook(f'{FileXLSX}.xlsx')        # example: xlrd.open_workbook('Book1.xlsx')
        sheet = book.sheet_by_index(0)

        for i in range(0,len(sheet.col(0))):     # len(sheet.col(0) is the size of the first column
                recieverName = sheet.row(i)[0].value   #row in excel = i + 1. So if 'i' = 0 corresponds to row 1. '[0]' corresponds to first column in row i
                recieverEmail = sheet.row(i)[1].value  #'[1]' corresponds to first column in row i
                subjectEmail = f'Hello {recieverName} {emailSubject}'
                bodyEmail = emailBody
                mimeMessage = MIMEMultipart()
                mimeMessage['to'] = recieverEmail
                mimeMessage['subject'] = subjectEmail
                mimeMessage.attach(MIMEText(bodyEmail, 'plain'))
                raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()
                message = service.users().messages().send(userId='me', body={'raw': raw_string}).execute()
                print(f'Message sent successfully to {recieverName}\n{message}\n')

def main():
        FileType = int(input('\nWhich file will you import from?\n\t(1) .txt\n\t(2) .xlsx\nEnter 1 or 2: '))
        FileName = input('Enter the name of your file (case sensitive): ')
        emailSubject = input('Enter the contents of your subject(will appear as "Hello (name of recipient) (your subject)"): ')
        emailBody = input('Enter the contents of your email: ')
        correct = input(f'\nIs this correct:\njson file: {credentials}\nFile: {FileName}\nSubject: Hello (name of recipient) {emailSubject}\nBody: {emailBody}\n\t1 for YES, 0 for NO\nEnter 1 or 0:')
        
        while(correct==0):
                FileType = int(input('\nWhich file will you import from?\n\t(1) .txt\n\t(2) .xlsx\nEnter 1 or 2: '))
                FileName = input('Enter the name of your file (case sensitive): ')
                emailSubject = input('Enter the contents of your subject(will appear as "Hello (name of recipient) (your subject)"): ')
                emailBody = input('Enter the contents of your email: ')
                correct = input(f'\nIs this correct:\njson file: {credentials}\nFile: {FileName}\nSubject: Hello (name of recipient) {emailSubject}\nBody: {emailBody}\n\t1 for YES, 0 for NO\nEnter 1 or 0:')
        if FileType == 1:
                usingTXT(FileTXT=FileName,emailBody=emailBody, emailSubject=emailSubject)
        if FileType == 2:
                usingXLSX(FileXLSX=FileName,emailBody=emailBody, emailSubject=emailSubject)

main()


