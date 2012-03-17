# -*- coding: utf-8 *-*
import gdata.spreadsheet.service
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
        self.khanTiersSheetIdMap = {}
        self.khanStudentsSheetIdMap = {}

        q = gdata.spreadsheet.service.DocumentQuery()
        q['title'] = self.khanTiersDoc
        q['title-exact'] = 'true'
        self.khanTiersFeed = self.client.GetSpreadsheetsFeed(query=q)
        self.khanTiersSpreadsheetId = \
            self.khanTiersFeed.entry[0].id.text.rsplit('/', 1)[1]
        self.khanTiersWorksheetFeed = \
            self.client.GetWorksheetsFeed(self.khanTiersSpreadsheetId)
        for w in self.khanTiersWorksheetFeed.entry:
            self.khanTiersSheetIdMap[w.content.text] = \
                w.id.text.rsplit('/', 1)[1]

        q['title'] = self.khanStudentsDoc
        q['title-exact'] = 'true'
        self.khanStudentsFeed = self.client.GetSpreadsheetsFeed(query=q)
        self.khanStudentsSpreadsheetId = \
            self.khanStudentsFeed.entry[0].id.text.rsplit('/', 1)[1]
        self.khanStudentsWorksheetFeed = \
            self.client.GetWorksheetsFeed(self.khanStudentsSpreadsheetId)
        for w in self.khanStudentsWorksheetFeed.entry:
            self.khanStudentsSheetIdMap[w.content.text] = \
                w.id.text.rsplit('/', 1)[1]


if __name__ == "__main__":

    GDC = GDocsClient()
    
    print GDC.khanTiersSheetIdMap
    print GDC.khanStudentsSheetIdMap
