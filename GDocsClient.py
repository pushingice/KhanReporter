# -*- coding: utf-8 *-*
import gdata.spreadsheet.service
import LoadStudents
import os.path

class GDocsClient():

    khanTiersDoc = "Khan Academy Tiers"
    khanStudentsDoc = "KhanTesting"

    def __init__(self):

        gdocPath = os.path.join("private","GDocs.txt")
        file = open(gdocPath).readlines()
        self.user = file[0].split()[1].strip()
        self.password = file[1].split()[1].strip()
        self.key = file[2].split()[1].strip()
        self.client = gdata.spreadsheet.service.SpreadsheetsService()
        self.client.ClientLogin(self.user, self.password)
        self.spreadsheetName = self.khanStudentsDoc
	self.spreadsheetKey = ""
	self.studentIDWorksheetName = "Khan Students"
	self.studentIDWorksheetKey = ""
#        self.khanTiersSheetIdMap = {}
#        self.khanStudentsSheetIdMap = {}
#        q = gdata.spreadsheet.service.DocumentQuery()
#        q['title'] = self.khanTiersDoc
#        q['title-exact'] = 'true'
#        self.khanTiersFeed = self.client.GetSpreadsheetsFeed(query=q)
#        self.khanTiersSpreadsheetId = \
#            self.khanTiersFeed.entry[0].id.text.rsplit('/', 1)[1]
#        self.khanTiersWorksheetFeed = \
#            self.client.GetWorksheetsFeed(self.khanTiersSpreadsheetId)
#        for w in self.khanTiersWorksheetFeed.entry:
#            self.khanTiersSheetIdMap[w.content.text] = \
#                w.id.text.rsplit('/', 1)[1]

#        q['title'] = self.khanStudentsDoc
#        q['title-exact'] = 'true'
#        self.khanStudentsFeed = self.client.GetSpreadsheetsFeed(query=q)
#        self.khanStudentsSpreadsheetId = \
#            self.khanStudentsFeed.entry[0].id.text.rsplit('/', 1)[1]
#        self.khanStudentsWorksheetFeed = \
#            self.client.GetWorksheetsFeed(self.khanStudentsSpreadsheetId)
#        for w in self.khanStudentsWorksheetFeed.entry:
#            self.khanStudentsSheetIdMap[w.content.text] = \
#                w.id.text.rsplit('/', 1)[1]


if __name__ == "__main__":

    GDC = GDocsClient()
    studentMap = LoadStudents.returnStudentMap()
    # Get the spreadsheet key
    feed = GDC.client.GetSpreadsheetsFeed()
    for f in feed.entry:
	if (f.title.text == GDC.spreadsheetName):
	    GDC.spreadsheetKey = f.id.text.split('/')[-1]

    # Get the worksheet key
    worksheetFeed = GDC.client.GetWorksheetsFeed(GDC.spreadsheetKey)
    for w in worksheetFeed.entry:
	if (w.title.text == GDC.studentIDWorksheetName):
	    GDC.studentIDWorksheetKey = w.id.text.split('/')[-1]
    
    #listFeed = GDC.client.GetListFeed(GDC.spreadsheetKey, GDC.studentIDWorksheetKey)
    #entry = GDC.client.UpdateCell(1,1,"chrisretford", GDC.spreadsheetKey, GDC.studentIDWorksheetKey)
    rowIndex = 2
    for k, v in studentMap.iteritems():
	print k, ' '.join(v)
	GDC.client.UpdateCell(rowIndex, 3, ' '.join(v), GDC.spreadsheetKey, GDC.studentIDWorksheetKey)
	GDC.client.UpdateCell(rowIndex, 4, k, GDC.spreadsheetKey, GDC.studentIDWorksheetKey)
	rowIndex += 1
	
