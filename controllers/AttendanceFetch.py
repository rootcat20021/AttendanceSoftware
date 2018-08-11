#Fetch incremental data
import os,time
import getpass, imaplib, email
import pandas as pd
import pickle

data_email_id = 'acknowledgesynchronization'
data_email_password = 'synchronizationacknowledge'
MailRate = 3600
M = ''
while(1):
    try:
        M = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    except:
        print "Cannot connect to gmail"
        time.sleep(MailRate)
        continue

    try:
        M.login(data_email_id+'@gmail.com',data_email_password)
    except:
        print "Unable to login to " + data_email_id + "@gmail.\n You probably need to open the email from your browser once or try to open:\n https://www.google.com/accounts/DisplayUnlockCaptcha"
        time.sleep(MailRate)
        continue

    M.select('inbox')
    result, data = M.uid('search', None, '(SUBJECT "SENDING_INCREMENTAL_UPDATE:")')
    uids = data[0].split()

    #Fetch the latest message
    try:
        result, data = M.uid('fetch', uids[-1], '(RFC822)')
    except:
        print(" client didn't acknowledged yet: ")
        time.sleep(MailRate)
        continue

    m = email.message_from_string(data[0][1])
    print "Subject = " + m['Subject']
    print m['Subject'].split(":")
    if m.get_content_maintype() == 'multipart': #multipart messages only
        for part in m.walk():
            if part.get_content_maintype() == 'multipart': continue
            if part.get('Content-Disposition') is None: continue

            #save the attachment in the program directory
            filename = part.get_filename()
            fp = open(os.path.join(filename), 'wb')
            fp.write(part.get_payload(decode=True))
            fp.close()
            print '%s saved!' % filename
            #Merge Attendance
            db.commit()
            pkl_file = open(filename, 'rb')
            AttendanceTodays = pickle.load(pkl_file)
            pathlog = os.path.join(request.folder,'private','log_auto_SSAttendance')
            logf = open(pathlog,'w')
            import commands
            logf.write(commands.getoutput('date'))
            logf.write("indexed SSAttendance Dates\n")
            logf.close()
            logf = open(pathlog,'a')

            LastUpdated = datetime.datetime(2000, 1, 1)
            datasource  = db(db.LocalVariables.id > 0).select()
            for data in datasource:
                LastUpdated = data['LastUpdated']

            logf.write('Attendance was last updated on ' + datetime.datetime.strftime(LastUpdated,'%d-%b-%Y') + '\n')

            #my headers are the headers in DB
            header_map_dict = {}
            header_map_dict['OldSewadarID'] =  'OldSewadarID'
            header_map_dict['SewadarNewID'] =  'SewadarNewID'
            header_map_dict['Name'] =  'Name'
            header_map_dict['Gender'] =  'Gender'
            header_map_dict['DepartmentID'] =  'DepartmentID'
            header_map_dict['DutyDate'] =  'DutyDate'
            header_map_dict['Duty Type'] =  'Duty_Type'

            for row in AttendanceTodays.iterrows():
               i=0
               row_dict = {}
               logf.write(str(row) + "\n")
               logf.close()
               logf = open(pathlog,'a')

               for col in header_map_dict.keys():
                   row_dict[header_map_dict[col]] = row[1][col]

               #print row_dict['SewadarNewID']
               if LastUpdated < row[1]['DutyDate']:
                   LastUpdated = row[1]['DutyDate']

               try:
                   db.SSAttendanceDate.insert(**row_dict)
               except:
                   db((db.SSAttendanceDate.SewadarNewID == row_dict['SewadarNewID']) & (db.SSAttendanceDate.DutyDate == row_dict['DutyDate'])).update(Duty_Type=row_dict['Duty_Type'])

    db(db.LocalVariables.id>0).update(LastUpdated=LastUpdated)
    logf.write("Send to database engine\n")

    db.commit()

    logf.write("Commited\n")
    logf.close()
    mail.send('acknowledgesynchronization@gmail.com',
        'Comitted SSAttendance to database',
        'Success',
        attachments = mail.Attachment('/home/rootcat/new_web2py/web2py/applications/AttendanceSoftware/private/log_auto_SSAttendance', content_id='text'))
    return 0


    time.sleep(MailRate)
