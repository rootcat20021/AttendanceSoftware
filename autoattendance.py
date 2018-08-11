import mechanize
import cookielib

# Browser
br = mechanize.Browser()

# Cookie Jar
cj = cookielib.LWPCookieJar()
br.set_cookiejar(cj)

# Browser options
br.set_handle_equiv(True)
br.set_handle_gzip(True)
br.set_handle_redirect(True)
br.set_handle_referer(True)
br.set_handle_robots(False)

# Follows refresh 0 but not hangs on refresh > 0
br.set_handle_refresh(mechanize._http.HTTPRefreshProcessor(), max_time=1)

# Want debugging messages?
#br.set_debug_http(True)
#br.set_debug_redirects(True)
#br.set_debug_responses(True)

# User-Agent (this is cheating, ok?)
br.addheaders = [('User-agent', 'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.1) Gecko/2008071615 Fedora/3.0.1-1.fc9 Firefox/3.0.1')]
# If the protected site didn't receive the authentication data you would
# end up with a 410 error in your face
br.add_password('https://110.172.171.195:13254/bimswebsite/Common/WebPages/Login.aspx', 'r', 'r')
r = br.open('https://110.172.171.195:13254/bimswebsite/Common/WebPages/Login.aspx')
#r = br.open('http://google.com')
html = r.read()

# Show the source
print html
# or
print br.response().read()

# Show the html title
print br.title()

# Show the response headers
print r.info()
# or
print br.response().info()

# Show the available forms
for f in br.forms():
    print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
    print f
    print "----------------------------"

# Select the first (index zero) form
br.select_form(nr=0)

# Let's search
br.form['txtPassword']='cafGA8338'
br.form['ctl00$ContentPlaceHolder1$txtUser']='g09700'
br.submit(name='ctl00$ContentPlaceHolder1$txtUser')
print br.response().read()

# Looking at some results in link format
for l in br.links(url_regex='stockrt'):
	    print l
