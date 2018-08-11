#Other database support
#db2 = DAL('mysql://rootcat:7133783@rootcat.mysql.pythonanywhere-services.com/rootcat$AttendanceDB',pool_size=1,check_reserved=['all'])
#db2 = DAL('mysql://sql695965:qT6!xZ4!@sql6.freesqldatabase.com:3306/sql695965',pool_size=1,check_reserved=['all'])
db2 = DAL('sqlite://storage8.db')

Incharge_Email_Ids = {}
Incharge_Email_Ids['1'] = 'badgecanteen1@gmail.com'
Incharge_Email_Ids['2'] = 'badgecanteen2@gmail.com'
Incharge_Email_Ids['3'] = 'badgecanteen3@gmail.com'
Incharge_Email_Ids['4'] = 'badgecanteen4@gmail.com'
Incharge_Email_Ids['5'] = 'badgecanteen5@gmail.com'
#Incharge_Email_Ids['1'] = 'pushkar.sareen@gmail.com'
#Incharge_Email_Ids['2'] = 'pushkar.sareen@gmail.com'
#Incharge_Email_Ids['3'] = 'pushkar.sareen@gmail.com'
#Incharge_Email_Ids['4'] = 'pushkar.sareen@gmail.com'
#Incharge_Email_Ids['5'] = 'pushkar.sareen@gmail.com'


from gluon.tools import Mail
mail = Mail()
mail.settings.server = 'smtp.gmail.com:25'
mail.settings.sender = 'acknowledgesynchronization@gmail.com'
mail.settings.login = 'acknowledgesynchronization:synchronizationacknowledge'

def Roundoffdate(DATETIME, CUTOFF):
    import datetime
    RoundedDate = DATETIME
    if (DATETIME.hour < int(CUTOFF)) or (DATETIME.hour == int(CUTOFF) and DATETIME.minute == 0) or (CUTOFF == 24):
       RoundedDate = RoundedDate.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
       RoundedDate = (RoundedDate + datetime.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    return RoundedDate

def uploaddata_SSAttendance():
    import xlrd
    import os
    import re, time
    import datetime
    import logging
    import pandas as pd
    db.commit()

    path = os.path.join(request.folder,'private','SSAttendanceDates_xls.xlsx')
    pathlog = os.path.join(request.folder,'private','log_SSAttendance')
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

    df = pd.read_excel(path,header=0,index_col=0)

    for row in df.iterrows():
       row_dict = {}
       #logf.write(str(row[1]) + "\n")
       #logf.close()
       #logf = open(pathlog,'a')

       row_dict = {}
       row_dict['SewadarNewID'] = row[1]['SewadarNewID']
       if type(row[1]['DutyDateTime']) == unicode:
           row_dict['DutyDate'] = datetime.datetime.strptime(row[1]['DutyDateTime'],'%d/%m/%Y %H:%M:%S')
           logf.write("unicode=")
       else:
           row_dict['DutyDate'] = row[1]['DutyDateTime'].replace(day=row[1]['DutyDateTime'].month,month=row[1]['DutyDateTime'].day)
           logf.write("datetime=")

       logf.write(datetime.datetime.strftime(row_dict['DutyDate'],'%d-%b-%Y') + '  ')
       logf.write("Diff = " + str(row_dict['DutyDate'].replace(hour=0, minute=0, second=0, microsecond=0) - row[1]['DutyDate'].replace(hour=0, minute=0, second=0, microsecond=0)) + '\n')
       row_dict['Duty_Type'] = row[1]['Duty Type']

       #print row_dict['SewadarNewID']
       if LastUpdated < row_dict['DutyDate']:
           LastUpdated = row_dict['DutyDate']

       try:
           db.SSAttendanceDate.insert(**row_dict)
       except:
           db((db.SSAttendanceDate.SewadarNewID == row_dict['SewadarNewID']) & (db.SSAttendanceDate.DutyDate == row_dict['DutyDate'])).update(Duty_Type=row_dict['Duty_Type'])

    db(db.LocalVariables.id>0).update(LastUpdated=LastUpdated)
    logf.write("Send to database engine\n")

    db.commit()

    logf.write("Commited\n")
    logf.close()
    mail.send('softwareattendance@gmail.com',
        'Comitted SSAttendance to database',
        'Success',
        attachments = mail.Attachment('/home/rootcat/new_web2py/web2py/applications/AttendanceSoftware/private/log_SSAttendance', content_id='text'))
    return 0


def ParshadListScheduled(DateSelectedStart,DateSelectedEnd,LALastLadiesNewGRNO,LBLastLadiesNewGRNO,GALastGentsNewGRNO,GBLastGentsNewGRNO,LastOSS,SSCountCutOffGents,SSCountCutOffLadies,CVCutOff,VisitCountCutOff,WWCutOff,WWWaiver,WWAgeWaiver,DAY_END_TIME,DumpMachineAttendance,DumpSSAttendance,MailSubject):
    db.commit()
    import os
    import pprint
    pathlog = os.path.join(request.folder,'private','log_ParshadListScheduled')
    logf = open(pathlog,'w')
    import os
    from openpyxl import Workbook
    from openpyxl.cell import get_column_letter
    from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
    response.headers['Connection'] =  'keep-alive'
    dworkbook = Workbook()
    dworkbookAttendance = Workbook()
    dpath = os.path.join(request.folder,'private','TentativeParshadList.xlsx')
    dpathAttendance = os.path.join(request.folder,'private','DumpAttendance.xlsx')


    from gluon.sqlhtml import form_factory
    import datetime
    import time


    logf.write("Commiting previous db changes\n")
    db.commit()
    logf.write("Committed\n")
    SSCount = db(db.SSAttendanceCount).select()
    logf.write("ffetched count\n")
    SSCountDict = {}
    for Sewadar in SSCount:
        SSCountDict[Sewadar.NewID,'Gender'] = Sewadar.gender
        SSCountDict[Sewadar.NewID,'TotalVisit'] = Sewadar.TotalVisit
        SSCountDict[Sewadar.NewID,'TotalCount'] = Sewadar.Total
        SSCountDict[Sewadar.NewID,'NAME'] = Sewadar.Name
        SSCountDict[Sewadar.NewID,'OldSewadarid'] = Sewadar.OldSewadarid
        SSCountDict[Sewadar.NewID,'status'] = Sewadar.status
        SSCountDict[Sewadar.NewID,'GENDER'] = Sewadar.gender
    #Next setup: Change current visit dates(move to old visit dates)
    #Change total visits

    pprint.pprint(SSCountDict,open('/home/rootcat/new_web2py/web2py/applications/AttendanceSoftware/private/log_SSCoundDict','wb'))



    SSCURRENT_VISIT_MORNING_START = {}
    SSCURRENT_VISIT_MORNING_START['D0'] = datetime.datetime.strptime('22-November-2017 00:00:00','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_MORNING_START['D1'] = datetime.datetime.strptime('23-November-2017 00:00:00','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_MORNING_START['D2'] = datetime.datetime.strptime('24-November-2017 00:00:00','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_MORNING_START['D3'] = datetime.datetime.strptime('25-November-2017 00:00:00','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_MORNING_START['D4'] = datetime.datetime.strptime('26-November-2017 00:00:00','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_MORNING_START['D5'] = datetime.datetime.strptime('27-November-2017 00:00:00','%d-%B-%Y %H:%M:%S')

    SSCURRENT_VISIT_MORNING_END = {}
    SSCURRENT_VISIT_MORNING_END['D0'] = datetime.datetime.strptime('22-November-2017 14:30:00','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_MORNING_END['D1'] = datetime.datetime.strptime('23-November-2017 14:30:00','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_MORNING_END['D2'] = datetime.datetime.strptime('24-November-2017 14:30:00','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_MORNING_END['D3'] = datetime.datetime.strptime('25-November-2017 14:30:00','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_MORNING_END['D4'] = datetime.datetime.strptime('26-November-2017 14:30:00','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_MORNING_END['D5'] = datetime.datetime.strptime('27-November-2017 14:30:00','%d-%B-%Y %H:%M:%S')

    SSCURRENT_VISIT_EVENING_START = {}
    SSCURRENT_VISIT_EVENING_START['D0'] = datetime.datetime.strptime('22-November-2017 14:30:01','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_EVENING_START['D1'] = datetime.datetime.strptime('23-November-2017 14:30:01','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_EVENING_START['D2'] = datetime.datetime.strptime('24-November-2017 14:30:01','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_EVENING_START['D3'] = datetime.datetime.strptime('25-November-2017 14:30:01','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_EVENING_START['D4'] = datetime.datetime.strptime('26-November-2017 14:30:01','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_EVENING_START['D5'] = datetime.datetime.strptime('27-November-2017 14:00:01','%d-%B-%Y %H:%M:%S')

    SSCURRENT_VISIT_EVENING_END = {}
    SSCURRENT_VISIT_EVENING_END['D0'] = datetime.datetime.strptime('22-November-2017 11:59:59','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_EVENING_END['D1'] = datetime.datetime.strptime('23-November-2017 11:59:59','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_EVENING_END['D2'] = datetime.datetime.strptime('24-November-2017 11:59:59','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_EVENING_END['D3'] = datetime.datetime.strptime('25-November-2017 11:59:59','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_EVENING_END['D4'] = datetime.datetime.strptime('26-November-2017 11:59:59','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT_EVENING_END['D5'] = datetime.datetime.strptime('27-November-2017 11:59:59','%d-%B-%Y %H:%M:%S')


    SSCURRENT_VISIT = {}
    SSCURRENT_VISIT['COUNT'] = 5
    SSCURRENT_VISIT['D0'] = datetime.datetime.strptime('22-November-2017 00:00:00','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT['D1'] = datetime.datetime.strptime('22-November-2017 12:00:00','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT['D2'] = datetime.datetime.strptime('23-November-2017 13:00:01','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT['D3'] = datetime.datetime.strptime('24-November-2017 09:00:01','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT['D4'] = datetime.datetime.strptime('25-November-2017 09:00:01','%d-%B-%Y %H:%M:%S')
    SSCURRENT_VISIT['D5'] = datetime.datetime.strptime('26-November-2017 09:00:01','%d-%B-%Y %H:%M:%S')


    VISIT_DATES = {}
    TOTAL_VISIT = 18
    VISIT_DATES['V0','COUNT'] = 5
    VISIT_DATES['V0','D0'] = datetime.datetime.strptime('05-October-2011 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V0','D1'] = datetime.datetime.strptime('05-October-2011 14:30:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V0','D2'] = datetime.datetime.strptime('06-October-2011 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V0','D3'] = datetime.datetime.strptime('07-October-2011 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V0','D4'] = datetime.datetime.strptime('08-October-2011 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V0','D5'] = datetime.datetime.strptime('09-October-2011 12:00:00','%d-%B-%Y %H:%M:%S')

    VISIT_DATES['V1','COUNT'] = 5
    VISIT_DATES['V1','D0'] = datetime.datetime.strptime('23-November-2011 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V1','D1'] = datetime.datetime.strptime('23-November-2011 14:30:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V1','D2'] = datetime.datetime.strptime('24-November-2011 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V1','D3'] = datetime.datetime.strptime('25-November-2011 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V1','D4'] = datetime.datetime.strptime('26-November-2011 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V1','D5'] = datetime.datetime.strptime('27-November-2011 12:00:00','%d-%B-%Y %H:%M:%S')

    VISIT_DATES['V2','COUNT'] = 5
    VISIT_DATES['V2','D0'] = datetime.datetime.strptime('07-March-2012 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V2','D1'] = datetime.datetime.strptime('07-March-2012 14:30:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V2','D2'] = datetime.datetime.strptime('08-March-2012 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V2','D3'] = datetime.datetime.strptime('09-March-2012 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V2','D4'] = datetime.datetime.strptime('10-March-2012 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V2','D5'] = datetime.datetime.strptime('11-March-2012 12:00:00','%d-%B-%Y %H:%M:%S')

    VISIT_DATES['V3','COUNT'] = 5
    VISIT_DATES['V3','D0'] = datetime.datetime.strptime('10-October-2012 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V3','D1'] = datetime.datetime.strptime('10-October-2012 14:30:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V3','D2'] = datetime.datetime.strptime('11-October-2012 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V3','D3'] = datetime.datetime.strptime('12-October-2012 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V3','D4'] = datetime.datetime.strptime('13-October-2012 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V3','D5'] = datetime.datetime.strptime('14-October-2012 12:00:00','%d-%B-%Y %H:%M:%S')

    VISIT_DATES['V4','COUNT'] = 5
    VISIT_DATES['V4','D0'] = datetime.datetime.strptime('28-November-2012 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V4','D1'] = datetime.datetime.strptime('28-November-2012 14:30:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V4','D2'] = datetime.datetime.strptime('29-November-2012 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V4','D3'] = datetime.datetime.strptime('30-November-2012 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V4','D4'] = datetime.datetime.strptime('01-December-2012 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V4','D5'] = datetime.datetime.strptime('02-December-2012 12:00:00','%d-%B-%Y %H:%M:%S')

    VISIT_DATES['V5','COUNT'] = 5
    VISIT_DATES['V5','D0'] = datetime.datetime.strptime('13-March-2013 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V5','D1'] = datetime.datetime.strptime('13-March-2013 14:30:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V5','D2'] = datetime.datetime.strptime('14-March-2013 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V5','D3'] = datetime.datetime.strptime('15-March-2013 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V5','D4'] = datetime.datetime.strptime('16-March-2013 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V5','D5'] = datetime.datetime.strptime('17-March-2013 12:00:00','%d-%B-%Y %H:%M:%S')

    VISIT_DATES['V6','COUNT'] = 5
    VISIT_DATES['V6','D0'] = datetime.datetime.strptime('09-October-2013 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V6','D1'] = datetime.datetime.strptime('09-October-2013 14:30:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V6','D2'] = datetime.datetime.strptime('10-October-2013 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V6','D3'] = datetime.datetime.strptime('11-October-2013 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V6','D4'] = datetime.datetime.strptime('12-October-2013 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V6','D5'] = datetime.datetime.strptime('13-October-2013 12:00:00','%d-%B-%Y %H:%M:%S')

    VISIT_DATES['V7','COUNT'] = 3
    VISIT_DATES['V7','D0'] = datetime.datetime.strptime('10-January-2014 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V7','D1'] = datetime.datetime.strptime('10-January-2014 14:30:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V7','D2'] = datetime.datetime.strptime('11-January-2014 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V7','D3'] = datetime.datetime.strptime('12-January-2014 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V7','D4'] = datetime.datetime.strptime('13-January-2014 12:00:00','%d-%B-%Y %H:%M:%S')

    VISIT_DATES['V8','COUNT'] = 5
    VISIT_DATES['V8','D0'] = datetime.datetime.strptime('01-October-2014 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V8','D1'] = datetime.datetime.strptime('01-October-2014 14:30:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V8','D2'] = datetime.datetime.strptime('02-October-2014 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V8','D3'] = datetime.datetime.strptime('03-October-2014 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V8','D4'] = datetime.datetime.strptime('04-October-2014 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V8','D5'] = datetime.datetime.strptime('05-October-2014 12:00:00','%d-%B-%Y %H:%M:%S')

    VISIT_DATES['V9','COUNT'] = 5
    VISIT_DATES['V9','D0'] = datetime.datetime.strptime('26-November-2014 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V9','D1'] = datetime.datetime.strptime('26-November-2014 14:30:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V9','D2'] = datetime.datetime.strptime('27-November-2014 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V9','D3'] = datetime.datetime.strptime('28-November-2014 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V9','D4'] = datetime.datetime.strptime('29-November-2014 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V9','D5'] = datetime.datetime.strptime('30-November-2014 12:00:00','%d-%B-%Y %H:%M:%S')

    VISIT_DATES['V10','COUNT'] = 5
    VISIT_DATES['V10','D0'] = datetime.datetime.strptime('11-March-2015 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V10','D1'] = datetime.datetime.strptime('11-March-2015 14:30:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V10','D2'] = datetime.datetime.strptime('12-March-2015 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V10','D3'] = datetime.datetime.strptime('13-March-2015 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V10','D4'] = datetime.datetime.strptime('14-March-2015 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V10','D5'] = datetime.datetime.strptime('15-March-2015 12:00:00','%d-%B-%Y %H:%M:%S')

    VISIT_DATES['V11','COUNT'] = 5
    VISIT_DATES['V11','D0'] = datetime.datetime.strptime('07-October-2015 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V11','D1'] = datetime.datetime.strptime('07-October-2015 14:30:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V11','D2'] = datetime.datetime.strptime('08-October-2015 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V11','D3'] = datetime.datetime.strptime('09-October-2015 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V11','D4'] = datetime.datetime.strptime('10-October-2015 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V11','D5'] = datetime.datetime.strptime('11-October-2015 12:00:00','%d-%B-%Y %H:%M:%S')


    VISIT_DATES['V12','COUNT'] = 5
    VISIT_DATES['V12','D0'] = datetime.datetime.strptime('25-November-2015 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V12','D1'] = datetime.datetime.strptime('25-November-2015 14:30:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V12','D2'] = datetime.datetime.strptime('26-November-2015 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V12','D3'] = datetime.datetime.strptime('27-November-2015 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V12','D4'] = datetime.datetime.strptime('28-November-2015 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V12','D5'] = datetime.datetime.strptime('29-November-2015 12:00:00','%d-%B-%Y %H:%M:%S')

    VISIT_DATES['V13','COUNT'] = 5
    VISIT_DATES['V13','D0'] = datetime.datetime.strptime('9-March-2016 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V13','D1'] = datetime.datetime.strptime('9-March-2016 14:30:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V13','D2'] = datetime.datetime.strptime('10-March-2016 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V13','D3'] = datetime.datetime.strptime('11-March-2016 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V13','D4'] = datetime.datetime.strptime('12-March-2016 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V13','D5'] = datetime.datetime.strptime('13-March-2016 12:00:00','%d-%B-%Y %H:%M:%S')

    VISIT_DATES['V14','COUNT'] = 5
    VISIT_DATES['V14','D0'] = datetime.datetime.strptime('28-September-2016 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V14','D1'] = datetime.datetime.strptime('28-September-2016 14:30:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V14','D2'] = datetime.datetime.strptime('29-September-2016 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V14','D3'] = datetime.datetime.strptime('30-September-2016 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V14','D4'] = datetime.datetime.strptime('01-October-2016 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V14','D5'] = datetime.datetime.strptime('02-October-2016 12:00:00','%d-%B-%Y %H:%M:%S')

    VISIT_DATES['V15','COUNT'] = 5
    VISIT_DATES['V15','D0'] = datetime.datetime.strptime('23-November-2016 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V15','D1'] = datetime.datetime.strptime('23-November-2016 14:30:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V15','D2'] = datetime.datetime.strptime('24-November-2016 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V15','D3'] = datetime.datetime.strptime('25-November-2016 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V15','D4'] = datetime.datetime.strptime('26-November-2016 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V15','D5'] = datetime.datetime.strptime('27-November-2016 12:00:00','%d-%B-%Y %H:%M:%S')


    VISIT_DATES['V16','COUNT'] = 5
    VISIT_DATES['V16','D0'] = datetime.datetime.strptime('08-March-2017 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V16','D1'] = datetime.datetime.strptime('08-March-2017 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V16','D2'] = datetime.datetime.strptime('09-March-2017 13:00:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V16','D3'] = datetime.datetime.strptime('10-March-2017 09:00:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V16','D4'] = datetime.datetime.strptime('11-March-2017 09:00:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V16','D5'] = datetime.datetime.strptime('12-March-2017 09:00:01','%d-%B-%Y %H:%M:%S')

    VISIT_DATES['V17','COUNT'] = 5
    VISIT_DATES['V17','D0'] = datetime.datetime.strptime('04-October-2017 00:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V17','D1'] = datetime.datetime.strptime('04-October-2017 12:00:00','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V17','D2'] = datetime.datetime.strptime('05-October-2017 13:00:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V17','D3'] = datetime.datetime.strptime('06-October-2017 09:00:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V17','D4'] = datetime.datetime.strptime('07-October-2017 09:00:01','%d-%B-%Y %H:%M:%S')
    VISIT_DATES['V17','D5'] = datetime.datetime.strptime('08-October-2017 09:00:01','%d-%B-%Y %H:%M:%S')


    message = "ALL OK "
    ParshadList = {}


    try:
        os.remove(dpath)
    except:
        pass

    try:
        os.remove(dpathAttendance)
    except:
        pass

    logf.write("Collecting SSDate\n")
    logf.write("SSDate @")
    logf.write(str(datetime.datetime.now()) + "\n")
    SSDate = db((db.SSAttendanceDate.DutyDate >= (datetime.datetime.strptime(DateSelectedStart,'%Y-%m-%d %H:%M:%S').replace(hour=0, minute=0, second=0, microsecond=0))) & (db.SSAttendanceDate.DutyDate <= (datetime.datetime.strptime(DateSelectedEnd,'%Y-%m-%d %H:%M:%S').replace(hour=23, minute=59, second=59, microsecond=999)))).select('SewadarNewID','DutyDate','Duty_Type')
    logf.write("MAchineDate @")
    logf.write(str(datetime.datetime.now()) + "\n")
    MachineDate = db((db.MachineAttendance.DATETIME >= (datetime.datetime.strptime(DateSelectedStart,'%Y-%m-%d %H:%M:%S') - datetime.timedelta(hours=((24 - int(DAY_END_TIME)) % 24)))) & (db.MachineAttendance.DATETIME <= (datetime.datetime.strptime(DateSelectedEnd,'%Y-%m-%d %H:%M:%S') + datetime.timedelta(hours=int(DAY_END_TIME))))).select('GRNO','NewGRNO','DATETIME','TYPE')
    logf.write("ExceptionMail @")
    logf.write(str(datetime.datetime.now()) + "\n")
    dExceptionMail = db(db.ParshadMailException).select()
    print "Machine Attendance between date"
    print "length of MachineDate =" + str(len(MachineDate))

    ExceptionMail = {}
    #Keywords for exception
    #ALL , Visits Count,Current Visit
    for row in dExceptionMail:
        ExceptionMail[row.NewGRNO,row.ExceptionField] = row.Status


    ParshadList['SEWADARS'] = []
    SSAttendanceDictionary = {}
    print "Collecting SSDate"
    logf.write("Collecting SSDate\n")
    for SSEntry in SSDate:
        ParshadList['SEWADARS'].append(SSEntry.SewadarNewID)
        if SSEntry.Duty_Type == 'W':
            logf.write("Going to pop any existing D type attendance as a SS WW is found for " + SSEntry.SewadarNewID + " " + str(SSEntry.DutyDate) + "\n")
            now = datetime.datetime.now()
            SSAttendanceDictionary[SSEntry.SewadarNewID,Roundoffdate(SSEntry.DutyDate,23),'W'] = SSEntry.DutyDate
            try:
                ParshadList[SSEntry.SewadarNewID,'SSCount'] = ParshadList[SSEntry.SewadarNewID,'SSCount'] + 2
            except:
                ParshadList[SSEntry.SewadarNewID,'SSCount'] = 2
            try:
                if SSEntry.DutyDate.year == now.year:
                    ParshadList[SSEntry.SewadarNewID,'SSWWCount'] = ParshadList[SSEntry.SewadarNewID,'SSWWCount'] + 1
            except:
                if SSEntry.DutyDate.year == now.year:
                    ParshadList[SSEntry.SewadarNewID,'SSWWCount'] = 1
            try:
                ParshadList[SSEntry.SewadarNewID,'SSWWCountOld'] = ParshadList[SSEntry.SewadarNewID,'SSWWCount']
            except:
                ParshadList[SSEntry.SewadarNewID,'SSWWCountOld'] = 0

            #A D type attendance was counted first if pop was successful. Hence need to decrement by 1 to compensate
            if SSAttendanceDictionary.pop((SSEntry.SewadarNewID,Roundoffdate(SSEntry.DutyDate,23),'D'),None) == None:
                pass
            else:
                ParshadList[SSEntry.SewadarNewID,'SSCount'] = ParshadList[SSEntry.SewadarNewID,'SSCount'] - 1
            logf.write("Popped to pop any existing D type attendance as a SS WW is found for " + SSEntry.SewadarNewID + " " + str(SSEntry.DutyDate) + "\n")
        else:
            try:
                a = SSAttendanceDictionary[SSEntry.SewadarNewID,Roundoffdate(SSEntry.DutyDate,23),'W']
            except:
                SSAttendanceDictionary[SSEntry.SewadarNewID,Roundoffdate(SSEntry.DutyDate,23),'D'] = SSEntry.DutyDate
                try:
                    ParshadList[SSEntry.SewadarNewID,'SSCount'] = ParshadList[SSEntry.SewadarNewID,'SSCount'] + 1
                except:
                    ParshadList[SSEntry.SewadarNewID,'SSCount'] = 1

        ParshadList[SSEntry.SewadarNewID,'SSCountOld'] = ParshadList[SSEntry.SewadarNewID,'SSCount']

    MachineAttendanceAdditional = {}


    try:
        db.tempMachineAttendanceAdditional.drop()
    except:
        pass

    #Now define the table
    db.define_table('tempMachineAttendanceAdditional',
            Field('GRNO','string'),
            Field('NewGRNO','string'),
            Field('Duty_Type','string'),
            Field('DutyDate','datetime'),
            Field('DutyDateList','list:string'),
            migrate=True,
            redefine=True,
            format='%(NewGRNO)s')

    try:
        db(db.tempMachineAttendanceAdditional.id > 0).delete()
    except:
        pass

    print "Collecting Machine Date"

    for MEntry in MachineDate:
        if ((MEntry.DATETIME < VISIT_DATES['V13','D5']) and (MEntry.DATETIME > VISIT_DATES['V13','D0'])) or ((MEntry.DATETIME < VISIT_DATES['V12','D5']) and (MEntry.DATETIME > VISIT_DATES['V12','D0'])) or ((MEntry.DATETIME < VISIT_DATES['V14','D5']) and (MEntry.DATETIME > VISIT_DATES['V14','D0'])):
            pass
        else:
            #Select earliest entry but give preference to WW attendance
            if MEntry.TYPE == 'WMANUAL':
                try:
                    a = SSAttendanceDictionary[MEntry.NewGRNO,Roundoffdate(MEntry.DATETIME,DAY_END_TIME),'W']
                except:
                    logf.write("Going to pop any existing D type attendance as a WW is found for " + MEntry.NewGRNO + " " + str(Roundoffdate(MEntry.DATETIME,DAY_END_TIME)) + "\n")
                    try:
                        MachineAttendanceAdditional[MEntry.NewGRNO,"",Roundoffdate(MEntry.DATETIME,DAY_END_TIME),'W'].append(MEntry.DATETIME)
                    except:
                        MachineAttendanceAdditional[MEntry.NewGRNO,"",Roundoffdate(MEntry.DATETIME,DAY_END_TIME),'W'] = [MEntry.DATETIME]
                    MachineAttendanceAdditional.pop((MEntry.NewGRNO,"",Roundoffdate(MEntry.DATETIME,DAY_END_TIME),'D'),None)
                    logf.write("Popped any existing D type attendance as a WW is found for " + MEntry.NewGRNO + " " + str(Roundoffdate(MEntry.DATETIME,DAY_END_TIME)) + "\n")
            else:
                try:
                    a = MachineAttendanceAdditional[MEntry.NewGRNO,"",Roundoffdate(MEntry.DATETIME,DAY_END_TIME),'W']
                except:
                    try:
                        a = SSAttendanceDictionary[MEntry.NewGRNO,Roundoffdate(MEntry.DATETIME,DAY_END_TIME),'W']
                    except:
                        try:
                            a = SSAttendanceDictionary[MEntry.NewGRNO,Roundoffdate(MEntry.DATETIME,DAY_END_TIME),'D']
                        except:
                            try:
                                MachineAttendanceAdditional[MEntry.NewGRNO,"",Roundoffdate(MEntry.DATETIME,DAY_END_TIME),'D'].append(MEntry.DATETIME)
                            except:
                                MachineAttendanceAdditional[MEntry.NewGRNO,"",Roundoffdate(MEntry.DATETIME,DAY_END_TIME),'D'] = [MEntry.DATETIME]



    print "Preparing SS Count"

    for key, value in MachineAttendanceAdditional.iteritems():
        NewGRNO, GRNO, DutyDate , DutyType= key
        ParshadList['SEWADARS'].append(NewGRNO)
        if DutyType == 'W':
            try:
                ParshadList[NewGRNO,'SSCount'] = ParshadList[NewGRNO,'SSCount'] + 2
            except:
                ParshadList[NewGRNO,'SSCount'] = 2

            #A D type attendance was counted first if pop was successful. Hence need to decrement by 1 to compensate
            try:
                a = SSAttendanceDictionary[NewGRNO,DutyDate,'D']
                ParshadList[NewGRNO,'SSCount'] = ParshadList[NewGRNO,'SSCount'] - 1
            except:
                pass

            if DutyDate.year == now.year:
                try:
                    ParshadList[NewGRNO,'SSWWCount'] = ParshadList[NewGRNO,'SSWWCount'] + 1
                except:
                    ParshadList[NewGRNO,'SSWWCount'] = 1
            logf.write("Found and incremented WMANUAL attendance : " + NewGRNO + " " + str(DutyDate) + "\n")
        else:
            try:
                ParshadList[NewGRNO,'SSCount'] = ParshadList[NewGRNO,'SSCount'] + 1
            except:
                ParshadList[NewGRNO,'SSCount'] = 1


    for key, value in MachineAttendanceAdditional.iteritems():
        MachineAttendanceAdditionalDictionary = {}
        NewGRNO, GRNO, DutyDate , DutyType = key
        ParshadList['SEWADARS'].append(NewGRNO)
        MachineAttendanceAdditionalDictionary['NewGRNO'] = NewGRNO
        MachineAttendanceAdditionalDictionary['GRNO'] = GRNO
        MachineAttendanceAdditionalDictionary['Duty_Type'] = DutyType
        MachineAttendanceAdditionalDictionary['DutyDate'] = DutyDate
        MachineAttendanceAdditionalDictionary['DutyDateList'] = value
        try:
            ParshadList[NewGRNO,'Before Visit Machine Additional'].append(DutyDate)
        except:
            ParshadList[NewGRNO,'Before Visit Machine Additional'] = [DutyDate]
        db.tempMachineAttendanceAdditional.insert(**MachineAttendanceAdditionalDictionary)


    del MachineAttendanceAdditional

    print "Creating Attendance Sheets"

    if DumpSSAttendance == 'YES':
        dSSdate = dworkbookAttendance.create_sheet(0)
        dSSdate.title = "SSDate"
        dSSdate.append(['SewadarNewID','DutyDate','Duty_Type'])
        for row in SSDate:
            dSSdate.append([row.SewadarNewID , row.DutyDate , row.Duty_Type])

    if DumpMachineAttendance == 'YES':
        dMachineAttendance = dworkbookAttendance.create_sheet(0)
        dMachineAttendance.title = "MachineAttendance"
        dMachineAttendance.append(['GRNO','NewGRNO','DATETIME','TYPE'])
        for row in MachineDate:
            dMachineAttendance.append([row.GRNO,row.NewGRNO,row.DATETIME,row.TYPE])

    MachineDifference = db(db.tempMachineAttendanceAdditional.id > 0).select()

    dMachineDifference = dworkbook.create_sheet(0)
    dMachineDifference.title = "MachineDifference"
    dMachineDifference.append(['NewGRNO','GRNO','Duty_Type','DutyDate','DutyDateList','TimeDifference'])
    for row in MachineDifference:
        DutyDateList = ", ".join(map(str, row.DutyDateList))
        n = len(row.DutyDateList)
        TimeDifference = max(map(datetime.datetime.strptime,row.DutyDateList,['%Y-%m-%d %H:%M:%S']*n)) - min(map(datetime.datetime.strptime,row.DutyDateList,['%Y-%m-%d %H:%M:%S']*n))
        dMachineDifference.append([row.NewGRNO,row.GRNO,row.Duty_Type,row.DutyDate,DutyDateList,str(TimeDifference)])


    dworkbookAttendance.save(dpathAttendance)


    throttle = 1

    for visit in xrange(0,TOTAL_VISIT):
        VCOUNT = 0
        VCOUNT = VISIT_DATES['V'+str(visit),'COUNT']

        vc = 'V'+str(visit)

        for day in xrange(0,VCOUNT):
            dc = 'D'+str(day)
            dc1 = 'D'+str(day+1)
            print "Analyzing " + vc + dc

            SSDate = db((db.SSAttendanceDate.DutyDate >= datetime.datetime(VISIT_DATES[vc,dc].year,VISIT_DATES[vc,dc].month,VISIT_DATES[vc,dc].day,0,0,0)) & (db.SSAttendanceDate.DutyDate <= datetime.datetime((VISIT_DATES[vc,dc]+datetime.timedelta(hours=24)).year,(VISIT_DATES[vc,dc]+datetime.timedelta(hours=24)).month,(VISIT_DATES[vc,dc]+datetime.timedelta(hours=24)).day,0,0,0))).select('SewadarNewID')
            MachineDate = db((db.MachineAttendance.DATETIME >= VISIT_DATES[vc,dc]) & (db.MachineAttendance.DATETIME <= VISIT_DATES[vc,dc1])).select('NewGRNO')
            print "length of SSDate =" + str(len(SSDate))
            print "length of MachineDate =" + str(len(MachineDate))

            for row in SSDate:
                ParshadList['SEWADARS'].append(row.SewadarNewID)
                try:
                    a = ParshadList[row.SewadarNewID,vc]
                except:
                    ParshadList[row.SewadarNewID,vc] = VCOUNT - day

            for row in MachineDate:
                ParshadList['SEWADARS'].append(row.NewGRNO)
                try:
                    if ParshadList[row.NewGRNO,vc] < VCOUNT - day :
                        ParshadList[row.NewGRNO,vc] = VCOUNT - day
                        try:
                            ParshadList[row.NewGRNO,'Additional Machine Visit Days'].append(vc + dc + ' onwards')
                        except:
                            ParshadList[row.NewGRNO,'Additional Machine Visit Days'] = [vc + dc + ' onwards']
                except:
                    ParshadList[row.NewGRNO,vc] = (VCOUNT - day)
                    try:
                        ParshadList[row.NewGRNO,'Additional Machine Visit Days'].append(vc + dc + ' onwards')
                    except:
                        ParshadList[row.NewGRNO,'Additional Machine Visit Days'] = [vc + dc + ' onwards']
            del SSDate
            del MachineDate


    print "reading tentative parshad list"
    SSTentativeParshadList = db(db.SSTentativeParshadList.id > 0).select('NewGRNO','Status')

    for row in SSTentativeParshadList:
        ParshadList['SEWADARS'].append(row.NewGRNO)
        ParshadList[row.NewGRNO,'SS Tentative Parshad Status'] = row.Status


    print "reading initiated list"
    InitiatedList = db(db.InitiatedList.id > 0).select('NewGRNO','Status')

    for row in InitiatedList:
        ParshadList['SEWADARS'].append(row.NewGRNO)
        ParshadList[row.NewGRNO,'Initiation Status'] = row.Status


    print "reading previous parshad list"
    PreviousParshadList = db(db.PreviousParshadList.id > 0).select('NewGRNO','Status')

    for row in PreviousParshadList:
        ParshadList['SEWADARS'].append(row.NewGRNO)
        ParshadList[row.NewGRNO,'Previous Visit Parshad Status'] = row.Status

    print "Analyzing Current Visit"
    logf.write("Analkysinz current visit\n")

    #time.sleep (1);
    for day in xrange(0,SSCURRENT_VISIT['COUNT']):
        dc = 'D'+str(day)
        dc1 = 'D'+str(day+1)
        print "Analyzing Current Visit day " + str(day)

        SSDate = db((db.SSAttendanceDate.DutyDate >= SSCURRENT_VISIT['D'+str(day)]) & (db.SSAttendanceDate.DutyDate <= SSCURRENT_VISIT['D'+str(day+1)])).select('SewadarNewID')
        print "Collecting SSDate"
        time.sleep (1);
        MachineDate = db((db.MachineAttendance.DATETIME >= SSCURRENT_VISIT['D'+str(day)]) & (db.MachineAttendance.DATETIME <= SSCURRENT_VISIT['D'+str(day+1)])).select('NewGRNO')
        MachineDateMorning = db((db.MachineAttendance.DATETIME >= SSCURRENT_VISIT_MORNING_START['D'+str(day)]) & (db.MachineAttendance.DATETIME <= SSCURRENT_VISIT_MORNING_END['D'+str(day+1)])).select('NewGRNO')
        MachineDateEvening = db((db.MachineAttendance.DATETIME >= SSCURRENT_VISIT_EVENING_START['D'+str(day)]) & (db.MachineAttendance.DATETIME <= SSCURRENT_VISIT_EVENING_END['D'+str(day+1)])).select('NewGRNO')
        print "Collected Machine and SSDate"

        for row in SSDate:
            ParshadList['SEWADARS'].append(row.SewadarNewID)
            try:
                a = ParshadList[row.SewadarNewID,'CV']
            except:
                ParshadList[row.SewadarNewID,'CV'] = SSCURRENT_VISIT['COUNT'] - day
                ParshadList[row.SewadarNewID,'CVOld'] = SSCURRENT_VISIT['COUNT'] - day

            ParshadList[row.SewadarNewID,'CV '+ dc] = 'P'


        print "Analyzed SSDate for current visit"

        for row in MachineDate:
            ParshadList['SEWADARS'].append(row.NewGRNO)
            try:
                a = ParshadList[row.NewGRNO,'CV']
            except:
                ParshadList[row.NewGRNO,'CV'] = SSCURRENT_VISIT['COUNT'] - day

            try:
                a = ParshadList[row.NewGRNO,'CV '+ dc]
            except:
                ParshadList[row.NewGRNO,'CV '+ dc] = 'C'

        for row in MachineDateMorning:
            ParshadList['SEWADARS'].append(row.NewGRNO)
            ParshadList[row.NewGRNO,'CVMORNING',day] = 'P'

        for row in MachineDateEvening:
            ParshadList['SEWADARS'].append(row.NewGRNO)
            ParshadList[row.NewGRNO,'CVMORNING',day] = 'P'

        print "Analyzed Machine Dates for current visit"

    ParshadList['SEWADARS'] = set(ParshadList['SEWADARS'])

    print "Creating Super Set Parshad Status"
    logf.write("Creating Super set\n")

    dParshadList = dworkbook.create_sheet(0)
    dParshadList.title = "ParshadList"
    dParshadList.append(['NewGRNO','SSCount','SSCountOld','CVCount','CVCountOld','MachineBeforeVisitAdditional','WW Count','CV D1','CV D2','CV D3','CV D4','CV D5','MachinePreviousVisitAddition','InitiationStatus','SS Tentative Parshad Status','Previous Visit Parshad Status','VISIT_COUNTS','SSWWCount','V0','V1','V2','V3','V4','V5','V6','V7','V8','V9','V10','V11','V12','V13','V14'])
    row_num = 1
    for Sewadar in ParshadList['SEWADARS']:
        #print "SuperSet Sewadar " + Sewadar
        try:
            SSCount = ParshadList[Sewadar,'SSCount']
        except:
            SSCount = 0

        try:
            SSWWCount = ParshadList[Sewadar,'SSWWCount']
        except:
            SSWWCount = 0

        try:
            SSWWCountOld = ParshadList[Sewadar,'SSWWCountOld']
        except:
            SSWWCountOld = 0

        try:
            SSCountOld = ParshadList[Sewadar,'SSCountOld']
        except:
            SSCountOld = 0
        try:
            CVCount = ParshadList[Sewadar,'CV']
        except:
            CVCount = 0
        try:
            CVCountOld = ParshadList[Sewadar,'CVOld']
        except:
            CVCountOld = 0
        try:
            CVD0 = ParshadList[Sewadar,'CV D0']
        except:
            CVD0 = 'A'
        try:
            CVD1 = ParshadList[Sewadar,'CV D1']
        except:
            CVD1 = 'A'
        try:
            CVD2 = ParshadList[Sewadar,'CV D2']
        except:
            CVD2 = 'A'
        try:
            CVD3 = ParshadList[Sewadar,'CV D3']
        except:
            CVD3 = 'A'
        try:
            CVD4 = ParshadList[Sewadar,'CV D4']
        except:
            CVD4 = 'A'
        try:
            MachineBeforeVisitAdditional = ", ".join(map(str, ParshadList[Sewadar,'Before Visit Machine Additional']))
        except:
            MachineBeforeVisitAdditional = ""
        try:
            MachinePreviousVisitAddition = ", ".join(map(str, ParshadList[Sewadar,'Additional Machine Visit Days']))
        except:
            MachinePreviousVisitAddition = ""
        try:
            InitiationStatus = ParshadList[Sewadar,'Initiation Status']
        except:
            InitiationStatus = "Not Found"
        try:
            PreviousVisitParshadStatus = ParshadList[Sewadar,'Previous Visit Parshad Status']
        except:
            PreviousVisitParshadStatus = "Not Found"
        try:
            SSTentativeParshadStatus = ParshadList[Sewadar,'SS Tentative Parshad Status']
        except:
            SSTentativeParshadStatus = "Not Found"

        dParshadList.append([Sewadar,SSCount,SSCountOld,CVCount,CVCountOld,MachineBeforeVisitAdditional,SSWWCount,CVD0,CVD1,CVD2,CVD3,CVD4,MachinePreviousVisitAddition,InitiationStatus,SSTentativeParshadStatus,PreviousVisitParshadStatus])

        row_num = row_num + 1

        dParshadList.cell(get_column_letter(17)+str(row_num)).value = SSWWCountOld
        visit_counts = 0
        for visit in xrange(0,TOTAL_VISIT):
            try:
                dParshadList.cell(get_column_letter(visit+19)+str(row_num)).value = ParshadList[Sewadar,'V'+str(visit)]
                if ParshadList[Sewadar,'V'+str(visit)] >= 3:
                    visit_counts = visit_counts + 1
            except:
                dParshadList.cell(get_column_letter(visit+19)+str(row_num)).value = 0

        dParshadList.cell(get_column_letter(18)+str(row_num)).value = visit_counts


    MasterSheet = db(db.MasterSheet.id > 0).select('SewadarNewID','GR_NO','NAME','CANTEEN','DEV_DTY','Age')
    dWorkingParshad = dworkbook.create_sheet(0)
    dWorkingParshad.title = "WorkingParshad"
    dWorkingParshad.append(['SewadarNewID','GR_NO','NAME','CANTEEN','DEV_DTY','InitiationStatus','SS Tentative Parshad Status','CANTEEN Parshad Status','CANTEEN Parshad Remarks','SSCount','SSCountNew','CVCount','CVCountNew','MachineBeforeVisitAdditional','WW Count','CV D1','CV D2','CV D3','CV D4','CV D5','MachinePreviousVisitAddition','Previous Visit Parshad Status','SSWWCOUNTOLD','Gender','Age','VISIT_COUNTS','V0','V1','V2','V3','V4','V5','V6','V7','V8','V9','V10','V11','V12','V13','V14'])

    print "Creating Master Specific Parshad Status"
    logf.write("Creating Master Specific Parshad Status\n")

    row_num = 2
    for row in MasterSheet:
        Sewadar = row.SewadarNewID
        #print "Master Sewadar " + Sewadar

        dWorkingParshad.cell(get_column_letter(1) + str(row_num)).value = row.SewadarNewID
        dWorkingParshad.cell(get_column_letter(2) + str(row_num)).value = row.GR_NO
        dWorkingParshad.cell(get_column_letter(3) + str(row_num)).value = row.NAME
        dWorkingParshad.cell(get_column_letter(4) + str(row_num)).value = row.CANTEEN
        dWorkingParshad.cell(get_column_letter(5) + str(row_num)).value = row.DEV_DTY

        Age = row.Age




        try:
            SSCount = ParshadList[Sewadar,'SSCount']
        except:
            SSCount = 0
        try:
            SSCountOld = ParshadList[Sewadar,'SSCountOld']
        except:
            SSCountOld = 0
        try:
            SSWWCount = ParshadList[Sewadar,'SSWWCount']
        except:
            SSWWCount = 0
        try:
            SSWWCountOld = ParshadList[Sewadar,'SSWWCountOld']
        except:
            SSWWCountOld = 0
        try:
            CVCount = ParshadList[Sewadar,'CV']
        except:
            CVCount = 0
        try:
            CVCountOld = ParshadList[Sewadar,'CVOld']
        except:
            CVCountOld = 0
        try:
            CVD0 = ParshadList[Sewadar,'CV D0']
        except:
            CVD0 = 'A'
        try:
            CVD1 = ParshadList[Sewadar,'CV D1']
        except:
            CVD1 = 'A'
        try:
            CVD2 = ParshadList[Sewadar,'CV D2']
        except:
            CVD2 = 'A'
        try:
            CVD3 = ParshadList[Sewadar,'CV D3']
        except:
            CVD3 = 'A'
        try:
            CVD4 = ParshadList[Sewadar,'CV D4']
        except:
            CVD4 = 'A'
        try:
            MachineBeforeVisitAdditional = ", ".join(map(str, ParshadList[Sewadar,'Before Visit Machine Additional']))
        except:
            MachineBeforeVisitAdditional = ""
        try:
            MachinePreviousVisitAddition = ", ".join(map(str, ParshadList[Sewadar,'Additional Machine Visit Days']))
        except:
            MachinePreviousVisitAddition = ""
        try:
            InitiationStatus = ParshadList[Sewadar,'Initiation Status']
        except:
            InitiationStatus = "Not Found"
        try:
            PreviousVisitParshadStatus = ParshadList[Sewadar,'Previous Visit Parshad Status']
        except:
            PreviousVisitParshadStatus = "Not Found"
        try:
            SSTentativeParshadStatus = ParshadList[Sewadar,'SS Tentative Parshad Status']
        except:
            SSTentativeParshadStatus = "Not Found"

        dWorkingParshad.cell(get_column_letter(23)+str(row_num)).value = SSWWCountOld

        visit_counts = 0
        for visit in xrange(0,TOTAL_VISIT):
            try:
                dWorkingParshad.cell(get_column_letter(visit+27)+str(row_num)).value = ParshadList[Sewadar,'V'+str(visit)]
                if ParshadList[Sewadar,'V'+str(visit)] >= 3:
                    visit_counts = visit_counts + 1
            except:
                dWorkingParshad.cell(get_column_letter(visit+27)+str(row_num)).value = 0

        try:
            dWorkingParshad.cell(get_column_letter(24)+str(row_num)).value = SSCountDict[Sewadar,'GENDER']
        except:
            dWorkingParshad.cell(get_column_letter(24)+str(row_num)).value = "Missing in SSCount sheet"

        try:
            dWorkingParshad.cell(get_column_letter(25)+str(row_num)).value = Age
        except:
            dWorkingParshad.cell(get_column_letter(25)+str(row_num)).value = "Missing in Master.."



        try:
            if visit_counts < SSCountDict[Sewadar,'TotalVisit']:
                visit_counts = SSCountDict[Sewadar,'TotalVisit']
        except:
            test = 1

        dWorkingParshad.cell(get_column_letter(26)+str(row_num)).value = visit_counts

        CanteenParshadStatus = "OK"
        CanteenParshadRemark = []


        try:
            if (SSCountDict[row.SewadarNewID,'status'].upper() == 'PERMANENT') | (SSCountDict[row.SewadarNewID,'status'].upper() == 'ELDERLY'):
                pass
            else:
                if (int(visit_counts) < int(VisitCountCutOff)):
                    try:
                        CanteenParshadRemark.append(ExceptionMail[row.SewadarNewID,'Visits Count'])
                        CanteenParshadStatus = "Tentative"
                    except:
                        try:
                            CanteenParshadRemark = [ExceptionMail[row.SewadarNewID,'ALL']]
                        except:
                            CanteenParshadRemark.append(str(visit_counts) + " Visits Attended")
                            CanteenParshadStatus = "Not OK"
        except:
            CanteenParshadRemark.append("SS Missing")
            CanteenParshadStatus = "SS Missing"

        if (CVCount < int(CVCutOff)):
            try:
                CanteenParshadRemark.append(ExceptionMail[row.SewadarNewID,'Current Visit'])
                if CanteenParshadStatus == "OK":
                    CanteenParshadStatus = "Tentative"
            except:
                try:
                    CanteenParshadRemark = [ExceptionMail[row.SewadarNewID,'ALL']]
                except:
                    CanteenParshadRemark.append("Current Visit Short")
                    CanteenParshadStatus = "Not OK"

        if (row.SewadarNewID.find("G") > -1):
            if (SSCount >= int(WWWaiver)) | (Age >= int(WWAgeWaiver)):
                pass
            else:
                logf.write(str(SSCount) + ' < ' + str(WWWaiver) + '\n')
                if (SSWWCount < int(WWCutOff)):
                    try:
                        CanteenParshadRemark.append(ExceptionMail[row.SewadarNewID,'WW Count'])
                        if CanteenParshadStatus == "OK":
                            CanteenParshadStatus = "Tentative"
                    except:
                        try:
                            CanteenParshadRemark = [ExceptionMail[row.SewadarNewID,'ALL']]
                        except:
                            CanteenParshadRemark.append(str(SSWWCount) + " WW done")
                            CanteenParshadStatus = "Not OK"

            if (SSCount < int(SSCountCutOffGents)):
                try:
                    CanteenParshadRemark.append(ExceptionMail[row.SewadarNewID,'SS Count'])
                    if CanteenParshadStatus == "OK":
                        CanteenParshadStatus = "Tentative"
                except:
                    try:
                        CanteenParshadRemark = [ExceptionMail[row.SewadarNewID,'ALL']]
                    except:
                        CanteenParshadRemark.append(str(int(SSCountCutOffGents) - SSCount) + " Short Before Visit")
                        CanteenParshadStatus = "Not OK"

        if (row.SewadarNewID.find("L") > -1):
            if (SSCount < int(SSCountCutOffLadies)):
                try:
                    CanteenParshadRemark.append(ExceptionMail[row.SewadarNewID,'SS Count'])
                    if CanteenParshadStatus == "OK":
                        CanteenParshadStatus = "Tentative"
                except:
                    try:
                        CanteenParshadRemark = [ExceptionMail[row.SewadarNewID,'ALL']]
                    except:
                        CanteenParshadRemark.append(str(int(SSCountCutOffLadies) - SSCount) + " Short Before Visit")
                        CanteenParshadStatus = "Not OK"

        if InitiationStatus.find("N") > -1:
            try:
                CanteenParshadRemark.append(ExceptionMail[row.SewadarNewID,'Initiation'])
                if CanteenParshadStatus == "OK":
                    CanteenParshadStatus = "Tentative"
            except:
                try:
                    CanteenParshadRemark = [ExceptionMail[row.SewadarNewID,'ALL']]
                except:
                    CanteenParshadRemark.append("Non Initiated")
                    CanteenParshadStatus = "Not OK"

        #This visit criteria for Parshad Packet is that the Sewadar should be present
        if (CVCount >= int(CVCutOff)) and CanteenParshadStatus == "Not OK":
            #CanteenParshadRemark = "PARSHAD PACKET"
            CanteenParshadStatus = "PACKET OK"

        if (SSTentativeParshadStatus == "Not Found") & ((CanteenParshadStatus == "OK") | (CanteenParshadStatus == "Tentative")):
            CanteenParshadStatus = "Waiting"

        dWorkingParshad.cell(get_column_letter(6 ) + str(row_num)).value = InitiationStatus
        dWorkingParshad.cell(get_column_letter(7 ) + str(row_num)).value = SSTentativeParshadStatus
        dWorkingParshad.cell(get_column_letter(8 ) + str(row_num)).value = CanteenParshadStatus
        if len(CanteenParshadRemark) == 0:
            dWorkingParshad.cell(get_column_letter(9 ) + str(row_num)).value = "OK"
        else:
            dWorkingParshad.cell(get_column_letter(9 ) + str(row_num)).value = "\n".join(CanteenParshadRemark)
        dWorkingParshad.cell(get_column_letter(10) + str(row_num)).value = SSCountOld
        dWorkingParshad.cell(get_column_letter(11) + str(row_num)).value = SSCount
        dWorkingParshad.cell(get_column_letter(12) + str(row_num)).value = CVCountOld
        dWorkingParshad.cell(get_column_letter(13) + str(row_num)).value = CVCount
        dWorkingParshad.cell(get_column_letter(14) + str(row_num)).value = MachineBeforeVisitAdditional
        dWorkingParshad.cell(get_column_letter(15) + str(row_num)).value = SSWWCount
        dWorkingParshad.cell(get_column_letter(16) + str(row_num)).value = CVD0
        dWorkingParshad.cell(get_column_letter(17) + str(row_num)).value = CVD1
        dWorkingParshad.cell(get_column_letter(18) + str(row_num)).value = CVD2
        dWorkingParshad.cell(get_column_letter(19) + str(row_num)).value = CVD3
        dWorkingParshad.cell(get_column_letter(20) + str(row_num)).value = CVD4
        dWorkingParshad.cell(get_column_letter(21) + str(row_num)).value = MachinePreviousVisitAddition
        dWorkingParshad.cell(get_column_letter(22) + str(row_num)).value = PreviousVisitParshadStatus


        row_num = row_num + 1

    print "Almost Done!!"
    logf.write("Almost done!\n")

    del ParshadList
    dworkbook.save(dpath)
    mail.send('softwareattendance@gmail.com',
        MailSubject,
        'Tentative Parshad List\n DateSelectedStart=' + str(DateSelectedStart) + '\n DateSelectedEnd=' + str(DateSelectedEnd) + '\n SSCountCutOffLadies=' + str(SSCountCutOffLadies) + '\n SSCountCutOffGents=' + str(SSCountCutOffGents) + '\n VisitCountCutOff=' + str(VisitCountCutOff) + '\n CVCutOff=' + str(CVCutOff) + '\n WWCutOff=' + str(WWCutOff)  + '\n WWWaiver=' + str(WWWaiver) + '\n WWAgeWaiver =' + str(WWAgeWaiver),
        attachments = mail.Attachment('/home/rootcat/new_web2py/web2py/applications/AttendanceSoftware/private/TentativeParshadList.xlsx', content_id='text'))

    logf.write("Mail sent!\n")
    logf.close()

    return dict(message=message)


#@auth.requires_login()
def AttendanceRegisterScheduledAll(DateSelectedStart,DateSelectedEnd):
    import os
    pathlog = os.path.join(request.folder,'private','log_AttendanceRegisterAll')
    logf = open(pathlog,'w')

    from gluon.sqlhtml import form_factory
    import datetime
    import time

    from openpyxl import Workbook
    from openpyxl.cell import get_column_letter
    import os
    response.headers['Connection'] =  'keep-alive'
    dworkbook = Workbook()
    dAttendanceRegister = dworkbook.create_sheet(0)
    dpath = os.path.join(request.folder,'private','AttendanceRegister.xlsx')

    Register = "EMPTY"
    DAttendanceRegisterTable = "EMPTY"
    datasource = {}
    columns = []
    orderby = []
    headers = {}
    excel_headers = {}
    DEV_DTY = ""
    TextMessage = "Jatha Wise Attendance Report"
    ReportDate =  0
    GENTS_REQUIRED = 30
    LADIES_REQUIRED = 36
    SS_GENTS_REQUIRED = 30
    SS_LADIES_REQUIRED = 36

    try:
        db.tempAttendanceRegisterTable.drop()
    except:
        pass

    SewadarDetails = {'Sewadars':[]}
#    form=form_factory(SQLField('JATHA','string',requires=IS_IN_DB(db,'MasterSheet.DEV_DTY','%(DEV_DTY)s')),SQLField('DateStart','date',default=datetime.datetime.today()-datetime.timedelta(days=31)),SQLField('DateEnd','date',default=datetime.datetime.today()),formname='DateSelect')
    my_list = db().select(db.MasterSheet.DEV_DTY, distinct=True).as_list()
    myjathalist = []
    for mydict in my_list:
        myjathalist.append(mydict['DEV_DTY'])

    myjathalist.sort()
    form=form_factory(SQLField('JATHA','string',requires=IS_IN_SET(myjathalist)),SQLField('DateStart','date',default=datetime.datetime.today()-datetime.timedelta(days=31)),SQLField('DateEnd','date',default=datetime.datetime.today()),SQLField('Download','string',requires=IS_IN_SET(['YES','NO']),default='NO'),formname='DateSelect')
    try:
        os.remove(dpath)
    except:
        pass

    download = request.vars.Download
    ReportDate = datetime.datetime.today()-datetime.timedelta(days=3001)

    #SSDate = db((db.SSAttendanceDate.DutyDate >= DateSelectedStart) & (db.SSAttendanceDate.DutyDate <= DateSelectedEnd) & (db.SSAttendanceDate.Duty_Type == "W")).select()
    SSDate = db((db.SSAttendanceDate.DutyDate >= DateSelectedStart) & (db.SSAttendanceDate.DutyDate <= DateSelectedEnd)).select()
    for Sewadar in SSDate:
        delta = Sewadar.DutyDate - datetime.datetime.strptime(DateSelectedStart,'%Y-%m-%d')
        SewadarDetails[Sewadar.SewadarNewID,'DAYS',delta.days] = Sewadar.DutyDate
        SewadarDetails[Sewadar.SewadarNewID,'TYPE',delta.days] = Sewadar.Duty_Type
        if ReportDate < Sewadar.DutyDate:
            ReportDate = Sewadar.DutyDate

    SSCount = db(db.SSAttendanceCount).select()
    for Sewadar in SSCount:
        SewadarDetails[Sewadar.NewID,'Gender'] = Sewadar.gender
        SewadarDetails[Sewadar.NewID,'TotalCount'] = Sewadar.Total
        SewadarDetails[Sewadar.NewID,'NAME'] = Sewadar.Name


    MasterTable = db(db.MasterSheet.GR_NO == 'SS015298').select()
    for Sewadar in MasterTable:
             SewadarDetails[Sewadar.GR_NO,'SewadarNewID'] = Sewadar.SewadarNewID
             SewadarDetails[Sewadar.SewadarNewID,'DEV_DTY'] = Sewadar.DEV_DTY

    DEV_DTY = request.vars.JATHA
#        if auth.user.username == 'admin':
#            DEV_DTY = request.vars.JATHA
#        else:
#            try:
#                DEV_DTY = SewadarDetails[auth.user.username,'DEV_DTY']
#            except:
#                #try:
#                    DEV_DTY = SewadarDetails[SewadarDetails[auth.user.username,'SewadarNewID'],'DEV_DTY']
#                #except:
#                #    TextMessage = 'This ID is not allocated to any Jatha'


    MasterTable = db(db.MasterSheet.id > 0).select()
    for Sewadar in MasterTable:
        try:
            #Check if this ID was in SewaSamiti list
            SewadarDetails[Sewadar.SewadarNewID,'NAME']
            SewadarDetails['Sewadars'].append(Sewadar.SewadarNewID)
            SewadarDetails[Sewadar.SewadarNewID,'CANTEEN'] = Sewadar.CANTEEN
            SewadarDetails[Sewadar.SewadarNewID,'DEV_DTY'] = Sewadar.DEV_DTY
            SewadarDetails[Sewadar.SewadarNewID,'OLD_ID'] = Sewadar.GR_NO
        except:
            TextMessage = 'Sewadar ID not in list'


    Fields = [Field('D%02i' %i,'string') for i in range((datetime.datetime.strptime(DateSelectedEnd,'%Y-%m-%d') - datetime.datetime.strptime(DateSelectedStart,'%Y-%m-%d')).days)]


    #Now define the table
    db.define_table('tempAttendanceRegisterTable',
            Field('GR_NO','string'),
            Field('SewadarNewID','string'),
            Field('NAME','string'),
            Field('DEV_DTY','string'),
            Field('CANTEEN','string'),
            Field('TOTAL','integer'),
            Field('STATUS','string'),
            Field('REQD','integer'),
            Field('GENDER','string'),
            Field('WWDATES','string'),
            *Fields,
            migrate=True,
            redefine=True,
            format='%(SewadarNewID)s')

    db(db.tempAttendanceRegisterTable.id > 0).delete()

    excel_headers = {1:'SewadarNewID',2:'GR_NO',3:'NAME',4:'DEV_DTY',5:'CANTEEN',6:'GENDER',7:'STATUS',8:'REQD',9:'WWDATES'}
    dAttendanceRegister.append(excel_headers.values())

    SewadarNumber = 1
    for Sewadar in SewadarDetails['Sewadars']:
        logf.write(Sewadar)
        mydict = {}
        SewadarNumber = SewadarNumber + 1
        mydict['SewadarNewID'] = Sewadar
        mydict['GR_NO'] = SewadarDetails[Sewadar,'OLD_ID']
        mydict['NAME'] = SewadarDetails[Sewadar,'NAME']
        mydict['DEV_DTY'] = SewadarDetails[Sewadar,'DEV_DTY']
        mydict['CANTEEN'] = SewadarDetails[Sewadar,'CANTEEN']
        mydict['TOTAL'] = SewadarDetails[Sewadar,'TotalCount']
        mydict['GENDER'] = SewadarDetails[Sewadar,'Gender']
        mydict['STATUS'] = ''
        mydict['REQD'] = (GENTS_REQUIRED - SewadarDetails[Sewadar,'TotalCount']) if (SewadarDetails[Sewadar,'Gender'] == 'Male') else (LADIES_REQUIRED - SewadarDetails[Sewadar,'TotalCount'])
        if mydict['REQD'] < 1:
            mydict['REQD'] = 0
        elif (((SewadarDetails[Sewadar,'Gender'] == 'Male')) and mydict['REQD'] > SS_GENTS_REQUIRED):
            mydict['REQD'] = SS_GENTS_REQUIRED
        elif (((SewadarDetails[Sewadar,'Gender'] == 'Female')) and mydict['REQD'] > SS_LADIES_REQUIRED):
            mydict['REQD'] = SS_LADIES_REQUIRED
        else:
            pass



        #Write to excel too
        dAttendanceRegister.cell('A' + str(SewadarNumber)).value = mydict['SewadarNewID']
        dAttendanceRegister.cell('B' + str(SewadarNumber)).value = mydict['GR_NO']
        dAttendanceRegister.cell('C' + str(SewadarNumber)).value = mydict['NAME']
        dAttendanceRegister.cell('D' + str(SewadarNumber)).value = mydict['DEV_DTY']
        dAttendanceRegister.cell('E' + str(SewadarNumber)).value = mydict['CANTEEN']
        dAttendanceRegister.cell('F' + str(SewadarNumber)).value = mydict['GENDER']
        dAttendanceRegister.cell('G' + str(SewadarNumber)).value = mydict['STATUS']
        dAttendanceRegister.cell('H' + str(SewadarNumber)).value = mydict['REQD']


        columns = ['tempAttendanceRegisterTable.SewadarNewID','tempAttendanceRegisterTable.GR_NO','tempAttendanceRegisterTable.NAME','tempAttendanceRegisterTable.DEV_DTY','tempAttendanceRegisterTable.CANTEEN','tempAttendanceRegisterTable.REQD']
        orderby=columns

        headers = {'tempAttendanceRegisterTable.SewadarNewID':{'label':T('SewadarNewID'),'class':'','width':12,'truncate':12,'selected': False},
           'tempAttendanceRegisterTable.GR_NO':{'label':T('GR_NO'),'class':'','width':10,'truncate':10,'selected': False},
           'tempAttendanceRegisterTable.NAME':{'label':T('NAME'),'class':'','width':10,'truncate': 10,'selected': False},
           'tempAttendanceRegisterTable.DEV_DTY':{'label':T('DEV_DTY'),'class':'','width':11,'truncate': 11,'selected': False},
           'tempAttendanceRegisterTable.CANTEEN':{'label':T('CANTEEN'),'class':'','width':13,'truncate': 13,'selected': False},
           'tempAttendanceRegisterTable.REQD':{'label':T('REQD'),'class':'','width':4,'truncate': 4,'selected': False}
           }


        for i in range((datetime.datetime.strptime(DateSelectedEnd,'%Y-%m-%d') - datetime.datetime.strptime(DateSelectedStart,'%Y-%m-%d')).days):
            dateindex = datetime.datetime.strptime(DateSelectedStart,'%Y-%m-%d') + datetime.timedelta(i)
            columns.append('tempAttendanceRegisterTable.D%02i' %i)
            headers['tempAttendanceRegisterTable.D%02i' %i] = {'label':datetime.date.strftime(dateindex,'%d\n%m'),'class':'','width':'3','truncate': 3,'selected': False}
            excel_headers[i+10] = (datetime.date.strftime(dateindex,'%d-%m'))
            logf.write(excel_headers[i+10])
            logf.write("\n")

            try:
                mydict['D%02i' %i] = SewadarDetails[Sewadar,'TYPE',i]
                if mydict['D%02i' %i] == 'W':
                    try:
                        mydict["WWDATES"] = mydict["WWDATES"] + '\n' +  excel_headers[i+10]
                    except:
                        mydict["WWDATES"] = excel_headers[i+10]
            except:
                mydict['D%02i' %i] = ''
            dAttendanceRegister.cell(get_column_letter(i+10) + str(SewadarNumber)).value = mydict['D%02i' %i]
            try:
                dAttendanceRegister.cell(get_column_letter(9) + str(SewadarNumber)).value = mydict["WWDATES"]
            except:
                pass

        db.tempAttendanceRegisterTable.insert(**mydict)

    for key in excel_headers.keys():
        dAttendanceRegister.cell(get_column_letter(key) + '1').value = excel_headers[key]

    Register = SQLTABLE(SSDate,headers='fieldname:capitalize')

    datasource = db(db.tempAttendanceRegisterTable).select(orderby=db.tempAttendanceRegisterTable.GR_NO)

    DAttendanceRegisterTable = SQLTABLE(datasource,columns=columns,headers=headers,orderby='GR_NO',_class='datatable')
    TextMessage = ""
    ReportDate = datetime.date.strftime(ReportDate,"%d-%m-%Y")
    dworkbook.save(dpath)
    if download == "YES":
        redirect(URL(r=request, f='download_AttendanceRegister'))
    #tables = plugins.powerTable
    #tables.datasource = datasource
    #tables.uitheme = 'cupertino'
    #tables.dtfeatures['sPaginationType'] = 'two button'
    #tables.keycolumn = 'tempAttendanceRegisterTable.id'
    #tables.showkeycolumn = False
    #tables.columns = ['tempAttendanceRegisterTable.GR_NO','tempAttendanceRegisterTable.DEV_DTY','tempAttendanceRegisterTable.CANTEEN']
    #tables.hiddencolumns = ['tempAttendanceRegisterTable.GR_NO','tempAttendanceRegisterTable.DEV_DTY','tempAttendanceRegisterTable.CANTEEN']
    #tables.headers = 'labels'
    #created_table = tables.create()
    #DAttendanceRegisterTable = plugin_powerTable(datasource)

    mail.send('softwareattendance@gmail.com',
        'All Attendance register',
        'All attendance register',
        attachments = mail.Attachment('/home/rootcat/new_web2py/web2py/applications/AttendanceSoftware/private/AttendanceRegister.xlsx', content_id='text'))

    logf.close()
    return 0

def AttendanceFetch():
    import os,time
    import getpass, imaplib, email
    import pandas as pd
    import numpy as np
    import pickle
    import pprint
    from StyleFrame import StyleFrame, Styler, utils

    sorting_order = {}
    sorting_order[0,1] = ['3','3','3']
    sorting_order[0,2] = ['3','3','3']
    sorting_order[0,3] = ['3','3','3']
    sorting_order[0,4] = ['3','3','3']
    sorting_order[0,5] = ['3','3','3']
    sorting_order[1,1] = ['2','2','2']
    sorting_order[1,2] = ['2','2','2']
    sorting_order[1,3] = ['2','2','2']
    sorting_order[1,4] = ['2','2','2']
    sorting_order[1,5] = ['2','2','2']
    sorting_order[2,1] = ['1','1','1']
    sorting_order[2,2] = ['2','2','2']
    sorting_order[2,3] = ['5','5','5']
    sorting_order[2,4] = ['4','4','4']
    sorting_order[2,5] = ['3','3','3']
    sorting_order[3,1] = ['4','4','4']
    sorting_order[3,2] = ['4','4','4']
    sorting_order[3,3] = ['4','4','4']
    sorting_order[3,4] = ['4','4','4']
    sorting_order[3,5] = ['4','4','4']
    sorting_order[4,1] = ['1','1','1']
    sorting_order[4,2] = ['1','1','1']
    sorting_order[4,3] = ['1','1','1']
    sorting_order[4,4] = ['1','1','1']
    sorting_order[4,5] = ['1','1','1']
    sorting_order[5,1] = ['5','5','5']
    sorting_order[5,2] = ['5','5','5']
    sorting_order[5,3] = ['5','5','5']
    sorting_order[5,4] = ['5','5','5']
    sorting_order[5,5] = ['5','5','5']
    sorting_order[6,1] = ['3','3','3']
    sorting_order[6,2] = ['4','4','4']
    sorting_order[6,3] = ['1','1','1']
    sorting_order[6,4] = ['2','2','2']
    sorting_order[6,5] = ['1','2','4']
    revolve = len(sorting_order[0,1])


    pathlog = os.path.join(request.folder,'private','log_auto_SSAttendance')
    logf = open(pathlog,'w')
    logf.write("starting\n")
    import commands
    logf.write(commands.getoutput('date'))

    #db.SSAttendanceDate.drop()
    data_email_id = 'acknowledgesynchronization'
    data_email_password = 'synchronizationacknowledge'
    M = ''
    try:
        M = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    except:
        logf.write("Cannot connect to gmail")

    try:
        M.login(data_email_id+'@gmail.com',data_email_password)
    except:
        logf.write("Unable to login to " + data_email_id + "@gmail.\n You probably need to open the email from your browser once or try to open:\n https://www.google.com/accounts/DisplayUnlockCaptcha")

    M.select('inbox')
    result, data = M.uid('search', None, '(SUBJECT "SENDING_INCREMENTAL_UPDATE:DATE: ")')
    uids = data[0].split()

    try:
        result, data = M.uid('fetch', uids[-1], '(RFC822)')
    except:
        logf.write(" client didn't acknowledged yet: ")

    m = email.message_from_string(data[0][1])
    if m.get_content_maintype() == 'multipart': #multipart messages only
        for part in m.walk():
            if part.get_content_maintype() == 'multipart': continue
            if part.get('Content-Disposition') is None: continue

            #save the attachment in the program directory
            filename = part.get_filename()
            fp = open(os.path.join(request.folder,'private',filename), 'wb')
            fp.write(part.get_payload(decode=True))
            fp.close()
            logf.write('%s saved!\n' % filename)


    result, data = M.uid('search', None, '(SUBJECT "SENDING_INCREMENTAL_UPDATE:COUNT: ")')
    uids = data[0].split()

    try:
        result, data = M.uid('fetch', uids[-1], '(RFC822)')
    except:
        logf.write(" Count client didn't acknowledged yet: ")

    m = email.message_from_string(data[0][1])
    if m.get_content_maintype() == 'multipart': #multipart messages only
        for part in m.walk():
            if part.get_content_maintype() == 'multipart': continue
            if part.get('Content-Disposition') is None: continue

            #save the attachment in the program directory
            filename = part.get_filename()
            fp = open(os.path.join(request.folder,'private',filename), 'wb')
            fp.write(part.get_payload(decode=True))
            fp.close()
            logf.write('%s saved!\n' % filename)




    #Update Count
    filename = 'attendance_count.pkl'
    db.commit()
    logf.write("Commited!\n")
    logf.write("loading attendance count pkl!\n")
    pkl_file = ""
    try:
        logf.write("loading pkl now!\n")
        pkl_file = open(os.path.join(request.folder,'private',filename), 'rb')
        logf.write("loaded pkl!\n")
    except:
        logf.write("Unable to open pkl\n")
    AttendanceCount = pd.read_pickle(os.path.join(request.folder,'private',filename))
    logf.write("read pickle!\n")
    db(db.SSAttendanceCount.id > 0).delete()


    #my headers are the headers in DB
    header_map_dict = {}
    header_map_dict['NewID'] =  'NewID'
    header_map_dict['OldSewadarid']        =  'OldSewadarid'
    header_map_dict['Name']                =  'Name'
    header_map_dict['Father_Husband_Name'] =  'Father_Husband_Name'
    header_map_dict['status']              =  'status'
    header_map_dict['Gender']              =  'gender'
    header_map_dict['B']                   =  'B'
    header_map_dict['w']                   =  'w'
    header_map_dict['V1']                  =  'V1'
    header_map_dict['V2']                  =  'V2'
    header_map_dict['V3']                  =  'V3'
    header_map_dict['V4']                  =  'V4'
    header_map_dict['TotalVisit']          =  'TotalVisit'
    header_map_dict['Total']               =  'Total'
    header_map_dict['areaname']            =  'areaname'

    for row in AttendanceCount.iterrows():
        i=0
        row_dict = {}
        logf.write(str(row) + "\n")
        logf.close()
        logf = open(pathlog,'a')

        for col in header_map_dict.keys():
            if(col == 'OldSewadarid'):
                if row[1][col] == 'NA' or row[1][col] == '' or pd.isnull(row[1][col]):
                    row_dict[header_map_dict[col]] = row[1]['NewID']
                else:
                    row_dict[header_map_dict[col]] = row[1][col]
            else:
               row_dict[header_map_dict[col]] = row[1][col]
           #logf.write(col + ' = ' + str(row[1][col]) + "\n")

        logf.write("abou to insert\n")
        try:
            db.SSAttendanceCount.insert(**row_dict)
            logf.write("insert succesfull\n")
        except:
            logf.write("Unable to update count\n")

    db.commit()

    ######################################################################
    #Merge Attendance
    ######################################################################
    filename = 'todays_attendance.pkl'
    db.commit()
    logf.write("Commited!\n")
    logf.write("loading pkl!\n")
    pkl_file = ""
    try:
        logf.write("loading pkl now!\n")
        pkl_file = open(os.path.join(request.folder,'private',filename), 'rb')
        logf.write("loaded pkl!\n")
    except:
        logf.write("Unable to open pkl\n")
    AttendanceTodays = pd.read_pickle(os.path.join(request.folder,'private',filename))
    logf.write("read pickle!\n")
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
    header_map_dict['SewadarNewID'] =  'SewadarNewID'
    header_map_dict['DutyDateTime'] =  'DutyDate'
    header_map_dict['Duty Type'] =  'Duty_Type'

    for row in AttendanceTodays.iterrows():
       i=0
       row_dict = {}
       logf.write(str(row) + "\n")
       #logf.close()
       #logf = open(pathlog,'a')

       for col in header_map_dict.keys():
           if(col == 'DutyDateTime'):
               row_dict['DutyDate'] = datetime.datetime.strptime(row[1]['DutyDateTime'],'%d/%m/%Y %H:%M:%S')
           else:
               row_dict[header_map_dict[col]] = row[1][col]
           #logf.write(col + ' = ' + str(row[1][col]) + "\n")

       logf.write("abou to insert\n")
       #print row_dict['SewadarNewID']
       if LastUpdated < row[1]['DutyDate']:
           LastUpdated = row[1]['DutyDate']

       logf.write("about to insert\n")
       try:
           db.SSAttendanceDate.insert(**row_dict)
           logf.write("insert succesfull\n")
       except:
           db((db.SSAttendanceDate.SewadarNewID == row_dict['SewadarNewID']) & (db.SSAttendanceDate.DutyDate == row_dict['DutyDate'])).update(Duty_Type=row_dict['Duty_Type'])
           logf.write("trying duty type updaate\n")

    db(db.LocalVariables.id>0).update(LastUpdated=LastUpdated)
    logf.write("Send to database engine\n")

    db.commit()

    logf.write("Commited\n")
    SSDates = {}
    if datetime.datetime.now().hour >= 19:
        SSDates = db((db.SSAttendanceDate.DutyDate >= (datetime.datetime.now().replace(hour=19, minute=0, second=0, microsecond=0)))).select().as_list()
    else:
        SSDates = db((db.SSAttendanceDate.DutyDate >= (datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)))).select().as_list()

    #SSDates = db(db.SSAttendanceDate).select().as_list()
    df_dates = pd.DataFrame.from_records(SSDates)
    df_dates.to_excel(os.path.join(request.folder,'private','df_dates.xlsx'))
    if len(df_dates.index) == 0:
        mail.send('canteenattendance@gmail.com',
            'Attendance Report :' + datetime.datetime.strftime(datetime.datetime.now(),'%d-%b-%Y'),
            'No Attendance marked in BIMS for today!')
        return 0
    df_Master = pd.DataFrame.from_records(db(db.MasterSheet).select().as_list())
    df_Master['SewadarNewID'] = df_Master['SewadarNewID'].apply(lambda x: "{}{}".format('BH0011',x))
    pprint.pprint(df_Master,stream=logf)
    df_count = pd.DataFrame.from_records(db(db.SSAttendanceCount).select().as_list())
    df_count.rename(columns={'NewID':'SewadarNewID'},inplace=True)

    logf.write('Going to join now\n')
    temp_df_daily_report = df_dates.merge(df_Master,on=['SewadarNewID'],how='left')
    temp_df_daily_report.to_excel(os.path.join(request.folder,'private','temp_df_daily_report.xlsx'))
    logf.write('Going to join again now\n')
    df_daily_report = temp_df_daily_report.merge(df_count,on=['SewadarNewID'],how='left')
    df_daily_report.rename(columns={'DEV_DTY':'JATHA'},inplace=True)
    df_daily_report.rename(columns={'SewadarNewID':'ID'},inplace=True)
    df_daily_report.sort_values(by=['CANTEEN','JATHA','ID'],axis=0,inplace=True)
    df_daily_report = df_daily_report.reset_index()
    pprint.pprint(df_daily_report,stream=logf)
    df_daily_report['ID'] = df_daily_report['ID'].str.replace('BH0011','')
    writer = pd.ExcelWriter(os.path.join(request.folder,'private','df_daily_report.xlsx'),engine='xlsxwriter')
    df_daily_report.to_excel(os.path.join(request.folder,'private','temppp_df_daily_report.xlsx'))
    df_daily_report_pivot = pd.pivot_table(df_daily_report,index=['CANTEEN'],columns=['gender'],values='ID',aggfunc='count',margins=True)
    weekday = datetime.datetime.today().weekday()
    day = datetime.datetime.today().day
    weeknumber = int((day-1)/7)+1

    ReportsLocalVariables = pd.DataFrame.from_records(db(db.ReportsLocalVariables.id > 0).select().as_list())
    ReportsLocalVariables = ReportsLocalVariables.set_index(['Weekday','Weeknumber'])

    if (ReportsLocalVariables.at[(weekday,weeknumber),'LastDate'].replace(hour=0, minute=0, second=0, microsecond=0) == datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)):
        ordernumber = ReportsLocalVariables.at[(weekday,weeknumber),'LastCanteenInchargeIndex']
    else:
        ordernumber = (ReportsLocalVariables.at[(weekday,weeknumber),'LastCanteenInchargeIndex'] + 1) % revolve

    db(db.ReportsLocalVariables.Weekday == weekday).update(LastCanteenInchargeIndex=ordernumber,LastDate=datetime.datetime.now(),Weekday=weekday)

    canteen_incharge = sorting_order[weekday,weeknumber][ordernumber]
    header_df = pd.DataFrame(columns=['CANTEEN DAILY REPORT: ' + datetime.datetime.today().strftime('%d-%b-%Y') ])
    header_df.set_value('INCHARGE','CANTEEN DAILY REPORT: ' + datetime.datetime.today().strftime('%d-%b-%Y'),'CANTEEN: ' + str(canteen_incharge))
    #df_daily_report_pivot = df_daily_report_pivot.append(df_daily_report_pivot.sum(numeric_only=True), ignore_index=True)
    #logf.write("rename:" + str(df_daily_report_pivot.index.values.tolist()[-1]) + '\n')
    #df_daily_report_pivot.rename(index={df_daily_report_pivot.index.values.tolist()[-1]:'TOTAL'},inplace=True)
    df_daily_report_pivot.to_excel(writer,sheet_name='Summary',startrow=3,startcol=0)
    header_df.to_excel(writer,sheet_name='Summary',startrow=0,startcol=0,index=True)
    df_daily_report_canteen_incharge = df_daily_report[df_daily_report['CANTEEN'] == str(canteen_incharge)]
    df_daily_report_canteen_incharge.index = range(1,len(df_daily_report_canteen_incharge)+1)
    df_daily_report_others = df_daily_report[df_daily_report['CANTEEN'] != str(canteen_incharge)]
    df_daily_report_others.index = range(1,len(df_daily_report_others)+1)
    df_daily_report_canteen_incharge.to_excel(writer,columns=['ID','Name','JATHA','CANTEEN','DutyDate'],freeze_panes=(1,1),sheet_name='Incharge',startrow=0,startcol=0)
    df_daily_report_others.to_excel(writer,columns=['ID','Name','JATHA','CANTEEN','DutyDate'],freeze_panes=(1,1),sheet_name='Others',startrow=0,startcol=0)
    workbook  = writer.book
    cell_format = workbook.add_format()
    cell_format.set_align('center')
    cell_format.set_border(1)

    cell_format1 = workbook.add_format()
    cell_format1.set_align('left')
    cell_format1.set_border(1)

    worksheet = writer.sheets['Summary']
    worksheet.set_column('B:B', 8, cell_format)
    worksheet.set_column('A:A', 8, cell_format)
    worksheet.set_column('C:C', 8, cell_format)
    worksheet.set_column('D:D', 8, cell_format)
    worksheet = writer.sheets['Incharge']
    worksheet.set_column('B:B', 8, cell_format)
    worksheet.set_column('A:A', 5, cell_format)
    worksheet.set_column('C:C', 28, cell_format1)
    worksheet.set_column('D:D', 28, cell_format1)
    worksheet.set_column('E:E', None, cell_format)
    worksheet.set_column('F:F', 18, cell_format)
    worksheet = writer.sheets['Others']
    worksheet.set_column('B:B', 8, cell_format)
    worksheet.set_column('A:A', 5, cell_format)
    worksheet.set_column('C:C', 28, cell_format1)
    worksheet.set_column('D:D', 28, cell_format1)
    worksheet.set_column('E:E', None, cell_format)
    worksheet.set_column('F:F', 18, cell_format)
    logf.write("Saving xls")
    writer.save()
    logf.write("Saved xls!")
    
    #writer = StyleFrame.ExcelWriter(os.path.join(request.folder,'private',"stylepandas.xlsx"))
    #sf=StyleFrame(df_daily_report)
    #sf.apply_column_style(cols_to_style=df_daily_report.columns, styler_obj=Styler(bg_color=utils.colors.white, bold=True, font=utils.fonts.arial,font_size=8),style_header=True)
    #sf.apply_headers_style(styler_obj=Styler(bg_color=utils.colors.blue, bold=True, font_size=8, font_color=utils.colors.white,number_format=utils.number_formats.general, protection=False))
    #writer.save()
    #logf.write(df_dates)
    logf.write("Tadaaaa!!")
    logf.close()
    mail.send('canteenattendance@gmail.com',
        'Attendance Report :' + datetime.datetime.strftime(datetime.datetime.now(),'%d-%b-%Y'),
        'Success',
        attachments = [mail.Attachment('/home/rootcat/new_web2py/web2py/applications/AttendanceSoftware/private/df_daily_report.xlsx', content_id='excel')])
    mail.send(Incharge_Email_Ids[canteen_incharge],
        'Attendance Report :' + datetime.datetime.strftime(datetime.datetime.now(),'%d-%b-%Y'),
        'Success',
        attachments = [mail.Attachment('/home/rootcat/new_web2py/web2py/applications/AttendanceSoftware/private/df_daily_report.xlsx', content_id='excel')])


    return 0

from gluon.scheduler import Scheduler
scheduler = Scheduler(db2)
import datetime
#scheduler.queue_task(AttendanceFetch,
#                    start_time=datetime.datetime.strptime('09-July-2018 11:00:00','%d-%B-%Y %H:%M:%S'),  # datetime
#                    stop_time=None,  # datetime
#                    timeout = 3600,  # seconds
#                    prevent_drift=True,
#                    period=3600*24,  # seconds
#                    immediate=False,
#                    repeats=0)
