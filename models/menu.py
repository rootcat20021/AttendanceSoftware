# -*- coding: utf-8 -*-
# this file is released under public domain and you can use without limitations

#########################################################################
## Customize your APP title, subtitle and menus here
#########################################################################

response.logo = A(B('web',SPAN(2),'py'),XML('&trade;&nbsp;'),
                  _class="brand",_href="http://www.web2py.com/")
response.title = request.application.replace('_',' ').title()
response.subtitle = ''

## read more at http://dev.w3.org/html5/markup/meta.name.html
response.meta.author = 'Your Name <you@example.com>'
response.meta.description = 'a cool new app'
response.meta.keywords = 'web2py, python, framework'
response.meta.generator = 'Web2py Web Framework'

## your http://google.com/analytics id
response.google_analytics_id = None

#########################################################################
## this is the main application menu add/remove items as required
#########################################################################

response.menu = [
    (T('Sewadar Details'), False, URL('default', 'view_sewadar'), []),
    (T('Jatha wise Attendance register'), False, URL('default', 'AttendanceRegister'), []),
    (SPAN('Admin',_class='highlighted'), False, URL('default', 'MyAdmin'),[
    (T('Update'), False, URL('default', 'update'), []),
    (T('Attendance Register Jatha Wise'), False, URL('default', 'AttendanceRegisterDetailed'), []),
    (T('Notice Board Short Attendance'), False, URL('default', 'ShortAttendance'), []),
    (T('Parshad List and machine difference'), False, URL('default', 'ParshadList'), []),
    (T('Jatha Wise Short Attendance Register'), False, URL('default', 'ShortAttendanceRegister'), []),
    (T('Attendance Register for entire Canteen'), False, URL('default', 'AttendanceRegisterAll'), []),
    (T('Proximity Card List Registration'), False, URL('default', 'CardList'), [])]),
]

DEVELOPMENT_MENU = True

#########################################################################
## provide shortcuts for development. remove in production
#########################################################################

def _():
    # shortcuts
    app = request.application
    ctr = request.controller
    # useful links to internal and external resources

if DEVELOPMENT_MENU: _()

if "auth" in locals(): auth.wikimenu()
