db = DAL('sqlite://storage.db')
#db2 = DAL('mysql://rootcat:7133783@rootcat.mysql.pythonanywhere-services.com/rootcat$AttendanceDB',pool_size=1,check_reserved=['all'])

db.define_table('LocalVariables',
                Field('LastUpdated','datetime'),
                migrate=True,
                redefine=True)
                
db.define_table('ReportsLocalVariables',
                Field('Weekday','integer'),
                Field('Weeknumber','integer'),
                Field('LastDate','datetime'),
                Field('LastCanteenInchargeIndex','integer'),
                migrate=True,
                redefine=True)

db.define_table('SSAttendanceDate',
                Field('SewadarNewID','string'),
                Field('DutyDate','datetime'),
                Field('Duty_Type','string'),
                primarykey=['SewadarNewID','DutyDate'],
                migrate=True,
                redefine=True)

db.SSAttendanceDate.SewadarNewID.requires = IS_NOT_EMPTY()
db.SSAttendanceDate.DutyDate.requires = IS_NOT_EMPTY()
db.SSAttendanceDate.Duty_Type.requires = IS_NOT_EMPTY()





db.define_table('SSAttendanceCount',
                Field('NewID','string'),
                Field('OldSewadarid','string'),
                Field('Name','string'),
                Field('Father_Husband_Name','string'),
                Field('status','string'),
                Field('gender','string'),
                Field('B','integer'),
                Field('w','integer'),
                Field('V1','integer'),
                Field('V2','integer'),
                Field('V3','integer'),
                Field('V4','integer'),
                Field('Initiated_Status','string'),
                Field('TotalVisit','integer'),
                Field('Total','integer'),
                Field('areaname','string'),
                migrate=True,
                redefine=True,
                format='%(SewadarNewID)s')

#Add checks here









db.define_table('MasterSheet',
                Field('SewadarNewID','string'),
                Field('CANTEEN','string'),
                Field('DEV_DTY','string'),
                Field('MOBILE','string'),
                migrate=True,
                redefine=True,
                format='%(SewadarNewID)s')

db.MasterSheet.SewadarNewID.requires = IS_NOT_EMPTY()
db.MasterSheet.CANTEEN.requires = IS_NOT_EMPTY()
db.MasterSheet.DEV_DTY.requires = IS_NOT_EMPTY()


db.define_table('MachineAttendance',
                Field('GRNO','string'),
                Field('NewGRNO','string'),
                Field('DATETIME','datetime'),
                Field('IO','string'),
                Field('TYPE','string'),
                primarykey=['NewGRNO','DATETIME'],
                redefine=True,
                migrate=True)


db.MachineAttendance.DATETIME.requires = IS_DATETIME()


db.define_table('AllCardList',
                Field('PROXIMITY_CARDNUMBER','string'),
                migrate=True,
                redefine=True,
                format='%(PROXIMITY_CARDNUMBER)s')


db.define_table('CardList',
                Field('SewadarNewID','string'),
                Field('PROXIMITY_CARDNUMBER','string'),
                migrate=True,
                redefine=True,
                format='%(SewadarNewID)s')

db.CardList.SewadarNewID.requires = IS_NOT_EMPTY()
db.CardList.PROXIMITY_CARDNUMBER.requires =  IS_MATCH('^00\d{8}')



db.define_table('PreviousParshadList',
                Field('NewGRNO','string'),
                Field('Status','string'),
                redefine=True,
                migrate=True)


#Clear every visit
db.define_table('ParshadMailException',
                Field('SewadarNewID','string'),
                Field('ExceptionField','string'),
                Field('Status','string'),
                redefine=True,
                migrate=True)

db.ParshadMailException.ExceptionField.requires = IS_IN_SET(['ALL','SS Count','WW Count','Current Visit','Visits Count','Initiation','Mandatory Days'])

db.define_table('SSTentativeParshadList',
                Field('SewadarNewID','string'),
                Field('SSTentativeStatus','string'),
                redefine=True,
                migrate=True)

db.define_table('VisitDates',
                Field('DateMorningStart','datetime'),
                Field('DateMorningEnd','datetime'),
                Field('DateEveningStart','datetime'),
                Field('DateEveningEnd','datetime'),
                Field('DateEnd','datetime'))

db.define_table('WWSchedule',
                Field('Jatha','string'),
                Field('Dates','date'),
                redefine=True,
                migrate=True)

db.define_table('WWScheduleLadies',
                Field('Jatha','string'),
                Field('Days','string'),
                redefine=True,
                migrate=True)

from gluon.tools import Auth
auth = Auth(db)
auth.settings.registration_requires_approval = True
auth.define_tables(username=True,signature=False)
