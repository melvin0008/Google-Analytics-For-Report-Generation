import googleanalytics as ga
import json
from docx import Document
from docx.shared import Inches


document = Document()
credentials = json.load(open('credentials.json'))
accounts = ga.authenticate(**credentials)
profile = accounts[0].webproperties[0].profile


document.add_heading('Monthly Report For February', 0)
"""
Basic Metrics
"""
document.add_heading('Basic Metrics', level=1)
# print "Basic Metrics"
query = profile.core.query.set(metrics=['ga:users','ga:sessions','ga:pageviews','ga:uniquePageviews','ga:pageviewsPerSession','ga:bounceRate','ga:percentNewSessions']).set('start_date', '2016-04-01').set({'end_date': '2016-04-15'})
rows = query.get().rows[0]
users = rows[0]
sessions = rows[1]
pageviews = rows[2]
uniquePageviews = rows[3]
pageviewsPerSession = rows[4]
bounceRate = rows[5]

paragraph = document.add_paragraph("Users: "+str(users)+"\n"+"Sessions: "+str(sessions)+"\n"+"Pageviews:"+ str(pageviews)+"\n"+"UniquePageviews: "+str(uniquePageviews)+"\n"+"PageviewsPerSession: "+str(pageviewsPerSession)+"\n"+"BounceRate: "+ str(bounceRate))
# print "Users: "+str(users)
# print "Sessions: "+str(sessions)
# print "Pageviews:"+ str(pageviews)
# print "UniquePageviews: "+str(uniquePageviews)
# print "PageviewsPerSession: "+str(pageviewsPerSession)
# print "BounceRate: "+ str(bounceRate)

"""
New and Returning Sessions
"""
document.add_heading('New and Returning Sessions', level=1)
# print "\n\n\nNew and Returning Sessions"
percentNewSessions = rows[6]
percentReturningSessions = 100-rows[6]
# print "PercentNewSessions"+ str(percentNewSessions)
# print "PercentReturningSessions"+str(percentReturningSessions)
paragraph1 = document.add_paragraph("PercentNewSessions: "+ str(percentNewSessions)+"\n"+"PercentReturningSessions: "+str(percentReturningSessions))



def print_table(noRows,noColumns,columnNames,rows):
    if len(rows)<noRows:
        noRows = len(rows)
    table = document.add_table(rows=noRows+1, cols=noColumns)
    table.style = 'TableGrid'
    hdr_cells = table.rows[0].cells
    for i in xrange(noColumns):
        hdr_cells[i].text = columnNames[i] 
    for i,item in enumerate(rows):
        for j in  xrange(noColumns):
            table.rows[i+1].cells[j].text = item[j] 

def get_correct_rows(query):
    values = query.get().rows 
    corrected_values=[]
    for i,row in enumerate(values):
        corrected_values_each=[str(i+1)+"."]
        for row_val in row:
            corrected_values_each.append(str(row_val))
        corrected_values.append(corrected_values_each) 
    return corrected_values

"""
Top 10 pages
"""
document.add_heading("Top 10 pages", level=1)
query = profile.core.query.metrics('ga:pageviews', 'ga:uniquePageviews').dimensions('ga:pagePath').set('start_date', '2016-04-01').set({'end_date': '2016-04-15'}).sort('ga:pageviews', descending=True).limit(10)
corrected_values = get_correct_rows(query)
print_table(10,4,["No.","Page Path","Pageviews","UniquePageviews"],corrected_values)

"""
Technology used to view the website
"""    
# print "\n\n\n Technology used to view the website"
document.add_heading("Technology used to view the website", level=1)
query = profile.core.query.metrics('ga:sessions').dimensions('ga:deviceCategory').set('start_date', '2016-04-01').set({'end_date': '2016-04-15'}).sort('ga:sessions', descending=True).limit(10)
corrected_values = get_correct_rows(query)
print_table(3,3,["No.","Device","Session"],corrected_values)


"""
Audience
"""
# print "\n\n\nAudience"
query = profile.core.query.metrics('ga:sessions').dimensions('ga:country').set('start_date', '2016-04-01').set({'end_date': '2016-04-15'}).sort('ga:sessions', descending=True)
rows = query.get().rows
numberOfCountries = len(rows)
document.add_heading("Audience", level=1)
paragraph1 = document.add_paragraph("Total Number of Countries : "+str(numberOfCountries))
corrected_values=[]
for i, country in enumerate(rows[:10]):
    corrected_values.append([str(i+1)+".", country[0],("%.2f" % (float(country[1]*100)/sessions))+"%"]) 
print_table(10,3,["No.","Country","% of Sessions"],corrected_values)
# print "Total number of countries :"+str(numberOfCountries) + "\n"

"""
Traffic Sources
"""
# print "\n\n\n Traffic Sources"
document.add_heading("Traffic Sources", level=1)
query = profile.core.query.metrics('ga:sessions').dimensions('ga:medium').set('start_date', '2016-04-01').set({'end_date': '2016-04-15'}).sort('ga:sessions', descending=True).limit(10)
corrected_values = get_correct_rows(query)
print_table(10,3,["No.","Source","Sessions"],corrected_values)

"""
Search Terms
"""
# print "\n\n\nSearch Terms"
document.add_heading("Search Terms", level=1)
query = profile.core.query.metrics('ga:sessions').dimensions('ga:keyword').set('start_date', '2016-04-01').set({'end_date': '2016-04-15'}).sort('ga:sessions', descending=True).limit(10)
corrected_values = get_correct_rows(query)
print_table(10,3,["No.","Search Term","Sessions"],corrected_values)


"""
Organic Sources
"""
# print "\n\n\nOrganic sources"
document.add_heading("Organic Sources", level=1)
query = profile.core.query.metrics('ga:sessions').dimensions('ga:source').filter(medium='organic').set('start_date', '2016-04-01').set({'end_date': '2016-04-15'}).limit(10).sort('ga:sessions', descending=True)
corrected_values = get_correct_rows(query)
print_table(10,3,["No.","Organic Sources","Sessions"],corrected_values)

"""
Referral Sources
"""
document.add_heading("Referral Sources", level=1)
query = profile.core.query.metrics('ga:sessions').dimensions('ga:source').filter(medium='referral').set('start_date', '2016-04-01').set({'end_date': '2016-04-15'}).limit(10).sort('ga:sessions', descending=True)
corrected_values = get_correct_rows(query)
print_table(10,3,["No.","Referral Source","Sessions"],corrected_values)

"""
Non Buffalo Referrals
"""
document.add_heading("Non Buffalo Referral Sources", level=1)
query = profile.core.query.metrics('ga:sessions').dimensions('ga:source').filter(medium='referral').filter(source__ncontains="buffalo").set('start_date', '2016-04-01').set({'end_date': '2016-04-15'}).limit(10).sort('ga:sessions', descending=True)
corrected_values = get_correct_rows(query)
print_table(10,3,["No.","Source","Sessions"],corrected_values)

"""
Top Pages Visited as a Result of Direct Traffic
"""
document.add_heading("Top Pages Visited as a Result of Direct Traffic", level=1)
query = profile.core.query.metrics('ga:sessions').dimensions('ga:pagePath').filter(medium='(none)').set('start_date', '2016-04-01').set({'end_date': '2016-04-15'}).sort('ga:sessions', descending=True).limit(10)
corrected_values = get_correct_rows(query)
print_table(10,3,["No.","Pages","Sessions"],corrected_values)


"""
Social Network
"""
document.add_heading("Social Network", level=1)
query = profile.core.query.metrics('ga:sessions').dimensions('ga:socialNetwork').set('start_date', '2016-04-01').set({'end_date': '2016-04-15'}).sort('ga:sessions', descending=True).limit(10)
corrected_values = get_correct_rows(query)
print_table(10,3,["No.","Social Networks","Sessions"],corrected_values)


document.save('demo.docx')