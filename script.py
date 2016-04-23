import googleanalytics as ga
import json
from docx import Document
from docx.shared import Inches



class GoogleAnalyticsReport:
    def __init__(self,startDate,endDate,heading):
        credentials = json.load(open('credentials.json'))
        accounts = ga.authenticate(**credentials)
        self.profile = accounts[0].webproperties[0].profile
        self.document = Document()
        self.document.add_heading(heading, 0)
        self.startDate = startDate
        self.endDate = endDate
        self.sessions = 0


    """
    Basic Metrics
    """
    def getBasicMetrics(self):
        self.document.add_heading('Basic Metrics', level=1)
        query = self.profile.core.query.set(metrics=['ga:users','ga:sessions','ga:pageviews','ga:uniquePageviews','ga:pageviewsPerSession','ga:bounceRate','ga:percentNewSessions']).set('start_date', self.startDate).set({'end_date': self.endDate})
        rows = query.get().rows[0]
        users = rows[0]
        self.sessions = rows[1]
        pageviews = rows[2]
        uniquePageviews = rows[3]
        pageviewsPerSession = rows[4]
        bounceRate = rows[5]
        paragraph = self.document.add_paragraph("Users: "+str(users)+"\n"+"Sessions: "+str(self.sessions)+"\n"+"Pageviews:"+ str(pageviews)+"\n"+"UniquePageviews: "+str(uniquePageviews)+"\n"+"PageviewsPerSession: "+str(pageviewsPerSession)+"\n"+"BounceRate: "+ str(bounceRate))
        
        """
        New and Returning Sessions
        """
        self.document.add_heading('New and Returning Sessions', level=1)
        percentNewSessions = rows[6]
        percentReturningSessions = 100-rows[6]
        paragraph1 = self.document.add_paragraph("PercentNewSessions: "+ str(percentNewSessions)+"\n"+"PercentReturningSessions: "+str(percentReturningSessions))


    """
    Function to Print Tables
    """
    def __print_table(self,noRows,noColumns,columnNames,rows):
        if len(rows)<noRows:
            noRows = len(rows)
        table = self.document.add_table(rows=noRows+1, cols=noColumns,style=self.document.styles['TableGrid'])
        hdr_cells = table.rows[0].cells
        for i in xrange(noColumns):
            hdr_cells[i].text = columnNames[i] 
        for i,item in enumerate(rows):
            for j in  xrange(noColumns):
                table.rows[i+1].cells[j].text = item[j] 

    """
    Function to get correct rows from the received query. Also adds column number and stringifies everything
    """
    def __get_correct_rows(self,query):
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
    def getTopTenPages(self):
        self.document.add_heading("Top 10 pages", level=1)
        query = self.profile.core.query.metrics('ga:pageviews', 'ga:uniquePageviews').dimensions('ga:pagePath').set('start_date', self.startDate).set({'end_date': self.endDate}).sort('ga:pageviews', descending=True).limit(10)
        corrected_values = self.__get_correct_rows(query)
        self.__print_table(10,4,["No.","Page Path","Pageviews","UniquePageviews"],corrected_values)

    """
    Technology used to view the website
    """
    def getTechUsedViewWebsite(self):    
        self.document.add_heading("Technology used to view the website", level=1)
        query = self.profile.core.query.metrics('ga:sessions').dimensions('ga:deviceCategory').set('start_date', self.startDate).set({'end_date': self.endDate}).sort('ga:sessions', descending=True).limit(10)
        corrected_values = self.__get_correct_rows(query)
        self.__print_table(3,3,["No.","Device","Session"],corrected_values)


    """
    Audience
    """
    def getAudience(self):
        query = self.profile.core.query.metrics('ga:sessions').dimensions('ga:country').set('start_date', self.startDate).set({'end_date': self.endDate}).sort('ga:sessions', descending=True)
        rows = query.get().rows
        numberOfCountries = len(rows)
        self.document.add_heading("Audience", level=1)
        paragraph1 = self.document.add_paragraph("Total Number of Countries : "+str(numberOfCountries))
        corrected_values=[]
        for i, country in enumerate(rows[:10]):
            corrected_values.append([str(i+1)+".", country[0],("%.2f" % (float(country[1]*100)/self.sessions))+"%"]) 
        self.__print_table(10,3,["No.","Country","% of Sessions"],corrected_values)


    """
    Traffic Sources
    """
    def getTrafficSources(self):
        self.document.add_heading("Traffic Sources", level=1)
        query = self.profile.core.query.metrics('ga:sessions').dimensions('ga:medium').set('start_date', self.startDate).set({'end_date': self.endDate}).sort('ga:sessions', descending=True).limit(10)
        corrected_values = self.__get_correct_rows(query)
        self.__print_table(10,3,["No.","Source","Sessions"],corrected_values)

    """
    Search Terms
    """
    def getSearchTerms(self):
        self.document.add_heading("Search Terms", level=1)
        query = self.profile.core.query.metrics('ga:sessions').dimensions('ga:keyword').set('start_date', self.startDate).set({'end_date': self.endDate}).sort('ga:sessions', descending=True).limit(10)
        corrected_values = self.__get_correct_rows(query)
        self.__print_table(10,3,["No.","Search Term","Sessions"],corrected_values)


    """
    Organic Sources
    """
    def getOrganicSources(self):
        self.document.add_heading("Organic Sources", level=1)
        query = self.profile.core.query.metrics('ga:sessions').dimensions('ga:source').filter(medium='organic').set('start_date', self.startDate).set({'end_date': self.endDate}).limit(10).sort('ga:sessions', descending=True)
        corrected_values = self.__get_correct_rows(query)
        self.__print_table(10,3,["No.","Organic Sources","Sessions"],corrected_values)

    """
    Referral Sources
    """
    def getReferralSources(self):
        self.document.add_heading("Referral Sources", level=1)
        query = self.profile.core.query.metrics('ga:sessions').dimensions('ga:source').filter(medium='referral').set('start_date', self.startDate).set({'end_date': self.endDate}).limit(10).sort('ga:sessions', descending=True)
        corrected_values = self.__get_correct_rows(query)
        self.__print_table(10,3,["No.","Referral Source","Sessions"],corrected_values)

    """
    Non Buffalo Referrals
    """
    def getNonBuffaloReferrals(self):
        self.document.add_heading("Non Buffalo Referral Sources", level=1)
        query = self.profile.core.query.metrics('ga:sessions').dimensions('ga:source').filter(medium='referral').filter(source__ncontains="buffalo").set('start_date', self.startDate).set({'end_date': self.endDate}).limit(10).sort('ga:sessions', descending=True)
        corrected_values = self.__get_correct_rows(query)
        self.__print_table(10,3,["No.","Source","Sessions"],corrected_values)

    """
    Top Pages Visited as a Result of Direct Traffic
    """
    def getTopPagesFromDirectTraffic(self):
        self.document.add_heading("Top Pages Visited as a Result of Direct Traffic", level=1)
        query = self.profile.core.query.metrics('ga:sessions').dimensions('ga:pagePath').filter(medium='(none)').set('start_date', self.startDate).set({'end_date': self.endDate}).sort('ga:sessions', descending=True).limit(10)
        corrected_values = self.__get_correct_rows(query)
        self.__print_table(10,3,["No.","Pages","Sessions"],corrected_values)


    """
    Social Network
    """
    def getSocialNetwork(self):
        self.document.add_heading("Social Network", level=1)
        query = self.profile.core.query.metrics('ga:sessions').dimensions('ga:socialNetwork').set('start_date', self.startDate).set({'end_date': self.endDate}).sort('ga:sessions', descending=True).limit(10)
        corrected_values = self.__get_correct_rows(query)
        self.__print_table(10,3,["No.","Social Networks","Sessions"],corrected_values)


    def buildDoc(self):
        self.getBasicMetrics()
        self.getTopTenPages()
        self.getTechUsedViewWebsite()
        self.getAudience()
        self.getTrafficSources()
        self.getSearchTerms()
        self.getOrganicSources()
        self.getReferralSources()
        self.getNonBuffaloReferrals()
        self.getTopPagesFromDirectTraffic()
        self.getTopPagesFromDirectTraffic()
        self.getSocialNetwork()
        self.document.save('demo.docx')


report = GoogleAnalyticsReport('2016-04-01','2016-04-15','Monthly Report for XYZ')
report.buildDoc()