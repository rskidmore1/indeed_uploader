from tika import parser
from stateAbrivs import * 
import smtplib
from mail import * 
from api import * 

#from email.parser import Parser
import imaplib
import base64
import os
import email 
import schedule 
import time 
import pdftotext


#parser = Parser()

bodytext = '' 

def upload(): 
    

    m = ''
    # Something in lines of http://stackoverflow.com/questions/348630/how-can-i-download-all-emails-with-attachments-from-gmail
    # Make sure you have IMAP enabled in your gmail settings.
    # Right now it won't download same file name twice even if their contents are different.

    import email
    import getpass, imaplib
    import sys

    detach_dir = '.'

    msg = ''
    
    if 'attachments' not in os.listdir(detach_dir):
        os.mkdir('attachments')
    
    

    try:
        imapSession = imaplib.IMAP4_SSL('imap.gmail.com')
        typ, accountDetails = imapSession.login(userName, passwd)
        if typ != 'OK':
            print 'Not able to sign in!'
            raise

                   
        files = os.listdir('.')
        theFileArray = [x for x in files if '.pdf' in x]
        for x in theFileArray: 
            os.remove(x)
            print 'Removed: ' + x 
        

        docxArray = [x for x in files if '.docx' in x]
        for x in docsArray:
            os.remove(x)
            print 'Removed: ' + x 
        #imapSession.select('(\\HasNoChildren)/Notuploaded') 
        imapSession.select('INBOX')
        #imapSession.select('INBOX')
        #print imapSession.list()

        typ, data = imapSession.search(None, 'ALL')
        if typ != 'OK':
            print 'Error searching Inbox.'
            raise
        
        # Iterating over all emails
        ids = data[0] # data is a list.
        id_list = ids.split() # ids is a space separated string 
        latest_email_id = id_list[-1] # get the latest
        print "Last email ID:"
        print latest_email_id

        typ, messageParts = imapSession.fetch(latest_email_id, '(RFC822)')
        if typ != 'OK':
            print 'Error fetching mail.'
            raise
        fileName = ''
        emailBody = messageParts[0][1]
        print emailBody
        
        mail = email.message_from_string(emailBody)
        for part in mail.walk():
            if part.get_content_maintype() == 'multipart':
                # print part.as_string() 
                continue
            if part.get('Content-Disposition') is None:
                # print part.as_string()
                continue
            fileName = part.get_filename()

            if bool(fileName):
                filePath = os.path.join(detach_dir, detach_dir, fileName)
                if not os.path.isfile(filePath) :
                    print fileName
                    fp = open(filePath, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()


        noAttachmentSubject = []
        for i in range(1, 30):
            typ, msg_data = imapSession.fetch(str(i), '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_string(response_part[1])
                    for header in [ 'subject', 'to', 'from' ]:
                        #print '%-8s: %s' % (header.upper(), msg[header])
                        value = '%-8s: %s' % (header.upper(), msg[header])
                        noAttachmentSubject.append(value)
        
        '''
        data = conn.fetch(latest_email_id, '(BODY[HEADER.FIELDS (SUBJECT FROM)])')
        header_data = data[1][0][1]
        parser = HeaderParser()
        msg = parser.parsestr(header_data)
        '''


        #result = imapSession.store(latest_email_id, '+X-GM-LABELS', 'Uploaded')
        #result2 = imapSession.store(latest_email_id, '-X-GM-LABELS', 'Notuploaded')
        result3 = imapSession.store(latest_email_id, '+X-GM-LABELS', '\\Trash')


        #result = m.store(emailid, '+FLAGS', '\\Deleted')
        #print "result: "
        #print result 
        #result = imapSession.store(latest_email_id, '+X-GM-LABELS', 'Uploaded')
        mov, data = imapSession.uid('STORE', latest_email_id , '+FLAGS', '(\\Deleted)')
        #result = imapSession.store(latest_email_id, '+FLAGS', '\\Deleted')



        imapSession.expunge()
        
        imapSession.close()
        imapSession.logout()
        
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()  
        server.login(user, password)

        sent_from1 = ''
        to1 = ['']  
        subject = "File name processed: "  + fileName #message from email above   
        text = noAttachmentSubject
        message1 = 'Subject: {}\n\n{}'.format(subject, text)

        server.sendmail(sent_from1, to1, message1)
        server.close()
        

    





     
        try: 
            files = os.listdir('.')

            theFileArray = [x for x in files if '.pdf' in x]

            print theFileArray


            theFile = theFileArray[0]

            print theFile

 
            # Load your PDF
            with open(theFile, "rb") as f:
                pdf = pdftotext.PDF(f)
            
            

            print pdf[0] 
            
            print ''
            print ''
            print ''


            #Comma possition 
            findCommaPosition = pdf[0].find(',')
            print 'Find Comma Var: '
            print findCommaPosition 


            #Find state
            print pdf[0][findCommaPosition - 20:findCommaPosition + 5]

            x = ''

            try:
                if('-' in pdf[0][findCommaPosition - 20:findCommaPosition + 5] ):
                    beforeDash, AfterDash = pdf[0][findCommaPosition - 20:findCommaPosition + 5].split('-')
                    AfterDash = AfterDash.strip()
                    print "BeforeDash: "
                    print beforeDash
                    print "AfterDash:"
                    print AfterDash
                    
                    x,y= AfterDash.splitlines()
                else:
                    x,y= pdf[0][findCommaPosition - 20:findCommaPosition + 5].splitlines() 
            except: 
                if('-' in pdf[0][findCommaPosition - 20:findCommaPosition + 5] ):
                    beforeDash, AfterDash = pdf[0][findCommaPosition - 20:findCommaPosition + 5].split('-')
                    AfterDash = AfterDash.strip()
                    print "BeforeDash: "
                    print beforeDash
                    print "AfterDash:"
                    print AfterDash
                    y = AfterDash
                else:     
                    y = pdf[0][findCommaPosition - 20:findCommaPosition + 5] 

            print x 
            print ''
            print y

            if ('\n' in y ):
                array = y.splitlines()
                print array
                sub = ','
                for text in array:
                    if sub in text:
                        b = text
            else: 
                b = y


            city, stateZip = b.split(',')
            print "stateZip: " + stateZip 
            space = ''
            stateAb = ''
            zip = ''
            state = ''
            try: 
                space, stateAb, zip = stateZip.split(' ')
            except: 
                space, stateAb = stateZip.split(' ')
            print "city: " + city 
            print "state: " + stateAb
            print "zip: " + zip 


            for st, ab in us_state_abbrev.items(): 
                if stateAb in ab: 
                    print "State from dict: " + st
                    state = st 
                    print "State in if else: " + state
                print "State after if else: " + state

            #name
            nameArray = theFile.split('_')
            print 'nameArray: '
            print nameArray
            subPDF = '.pdf'
            for text in nameArray:
                if subPDF in text:
                    namePart, dotPDF = text.split('.p')
                    print namePart
                    print dotPDF
                    #print len(nameArray)
                    nameArray = nameArray[:-1]
                    print nameArray
                    nameArray.append(namePart)
                    print nameArray
            
            firstName = nameArray[0]
            lastNameArray = nameArray[1:2]
            lastName = ' '.join(lastNameArray)

            print 'firstName: '
            print firstName
            print 'lastName: '
            print lastName 

            #Company 
            company = ' '.join(nameArray)


            #Email 

            findatSignPosition = pdf[0].find('@')

            print findatSignPosition 



            findatSignArray = pdf[0][findatSignPosition - 30:findatSignPosition + 20].splitlines()
            print findatSignArray

            emailFromArray =  [x for x in findatSignArray if '@' in x]
            emailFromPDF = emailFromArray[0]
            print "Email from PDF: " + emailFromPDF
            print ''
            print ''


            findatSignForPhoneArray = pdf[0][findatSignPosition:findatSignPosition + 40].splitlines()

            print 'Find pheon array: '
            print findatSignForPhoneArray
            number = findatSignForPhoneArray[1]
            print 'Phone: '
            print number

            #Record ID Type
            LeadSource = 'Indeed'
            print "Uploading to salesforce."

			#uploads
            leadupload = sf.Lead.create({'FirstName': firstName, 'LastName': lastName, 'Company' : company, 'Phone' : number, 'State' : state,  'City'  : city, 'Email' : emailFromPDF, 'RecordTypeId' : '', 'LeadSource': LeadSource})

            removeFile = os.remove(theFile)
            time.sleep(5)
            



        except: 
            
            server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            server.ehlo()  
            server.login(user, password)

            sent_from = ''
            to = ['']  
            subject = "Partner Upload encountered error for: " + theFile #message from email above   
            text = 'Hey, test email for upload except ' 
            message = 'Subject: {}\n\n{}'.format(subject, text)

            server.sendmail(sent_from, to, message)
            server.close()
            
            removeFile = os.remove(theFile)
            print "Exception from pdf split."

            
    except :
        print 'No more emails. or email error'


#upload()


otherCounter = 0 


schedule.every(10).minutes.do(upload) #Change from seconds to minutes

while 1: 
  schedule.run_pending()

