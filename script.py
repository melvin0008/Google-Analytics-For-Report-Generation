import googleanalytics as ga
import json

credentials = json.load(open('credentials.json'))
accounts = ga.authenticate(**credentials)
profile = accounts[0].webproperties[0].profile
# query = profile.core.query.set(metrics=['ga:users','ga:sessions','ga:pageviews','ga:uniquePageviews','ga:pageviewsPerSession','ga:bounceRate','ga:percentNewSessions']).set('start_date', '2016-02-01').set({'end_date': '2016-02-29'})
# rows = query.get().rows[0]
# users = rows[0]
# sessions = rows[1]
# pageviews = rows[2]
# uniquePageviews = rows[3]
# pageviewsPerSession = rows[4]
# bounceRate = rows[5]
# percentNewSessions = rows[6]
# percentReturningSessions = 100-rows[6]
# print "Users: "+str(users)
# print "Sessions: "+str(sessions)
# print "Pageviews:"+ str(pageviews)
# print "UniquePageviews: "+str(uniquePageviews)
# print "PageviewsPerSession: "+str(pageviewsPerSession)
# print "BounceRate: "+ str(bounceRate)
# print "PercentNewSessions"+ str(percentNewSessions)
# print "PercentReturningSessions"+str(percentReturningSessions)
# query = profile.core.query.metrics('ga:pageviews', 'ga:uniquePageviews').dimensions('ga:pagePath').set('start_date', '2016-02-01').set({'end_date': '2016-02-29'}).sort('ga:pageviews', descending=True).limit(10)
# rows = query.get().rows
# for i,row in enumerate(rows):
#     print str(i+1)+". "+row[0]+"   "+str(row[1])+"   "+str(row[2]) 
# query = profile.core.query.metrics('ga:sessions').dimensions('ga:deviceCategory').set('start_date', '2016-02-01').set({'end_date': '2016-02-29'}).sort('ga:sessions', descending=True)
# rows = query.get().rows
# for i,mobile in enumerate(rows):
#     print str(i+1)+". "+mobile[0]+"  "+str(mobile[1])
# query = profile.core.query.metrics('ga:sessions').dimensions('ga:country').set('start_date', '2016-02-01').set({'end_date': '2016-02-29'}).sort('ga:sessions', descending=True)
# rows = query.get().rows
# numberOfCountries = len(rows)
# for i, country in enumerate(rows[:10]):
#     print str(i) + country[0]+("%.2f" % (float(country[1]*100)/sessions))+"%"  
# Keywords
# query = profile.core.query.metrics('ga:sessions').dimensions('ga:keyword').set('start_date', '2016-02-01').set({'end_date': '2016-02-29'}).sort('ga:sessions', descending=True).limit(10)
# rows = query.get().rows
# print rows[:10]
# Traffic Sources
# query = profile.core.query.metrics('ga:sessions').dimensions('ga:medium').set('start_date', '2016-02-01').set({'end_date': '2016-02-29'}).sort('ga:sessions', descending=True)
# rows = query.get().rows
# print rows
# Social Network
# query = profile.core.query.metrics('ga:sessions').dimensions('ga:socialNetwork').set('start_date', '2016-02-01').set({'end_date': '2016-02-29'}).sort('ga:sessions', descending=True)
# rows = query.get().rows
# print rows
# organic sources
# query = profile.core.query.metrics('ga:sessions').dimensions('ga:source').filter(medium='organic').set('start_date', '2016-02-01').set({'end_date': '2016-02-29'}).sort('ga:sessions', descending=True)
# rows = query.get().rows
# print rows
# referral sources
# query = profile.core.query.metrics('ga:sessions').dimensions('ga:source').filter(medium='referral').set('start_date', '2016-02-01').set({'end_date': '2016-02-29'}).limit(10).sort('ga:sessions', descending=True)
# rows = query.get().rows
# print rows
# Referrals
# query = profile.core.query.metrics('ga:sessions').dimensions('ga:source').filter(medium='referral').filter(source__ncontains="buffalo").set('start_date', '2016-02-01').set({'end_date': '2016-02-29'}).limit(10).sort('ga:sessions', descending=True)
# rows = query.get().rows
# print rows
# Non buffalo Referrals 
# query = profile.core.query.metrics('ga:sessions').dimensions('ga:source').filter(medium='none').filter(source__ncontains="buffalo").set('start_date', '2016-02-01').set({'end_date': '2016-02-29'}).limit(10).sort('ga:sessions', descending=True)
# rows = query.get().rows
# print rows

query = profile.core.query.metrics('ga:sessions').dimensions('ga:pagePath').filter(medium='(none)').set('start_date', '2016-02-01').set({'end_date': '2016-02-29'}).sort('ga:sessions', descending=True).limit(10)
rows = query.get().rows
print rows
# for i,row in enumerate(rows):
#     print str(i+1)+". "+row[0]+"   "+str(row[1])+"   "+str(row[2])