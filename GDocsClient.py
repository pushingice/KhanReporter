# -*- coding: utf-8 *-*
import gdata.spreadsheet.service as service
import os.path
import json
import KhanOAuthConnector as KAOC

class GDocsClient():

    khanStudentsDoc = "KhanTesting"

    def __init__(self):

        gdocPath = os.path.join("private","GDocs.txt")
        file = open(gdocPath).readlines()
        self.user = file[0].split()[1].strip()
        self.password = file[1].split()[1].strip()
        self.key = file[2].split()[1].strip()
        self.client = service.SpreadsheetsService()
        self.client.ClientLogin(self.user, self.password)
        self.spreadsheetName = self.khanStudentsDoc
	self.spreadsheetKey = ""
	self.studentIDWorksheetName = "aa_KhanStudents"
	self.studentIDWorksheetKey = ""

	self.khanExercisesWorksheetName = "aa_KhanExercises"
	self.khanExercisesWorksheetKey = ""

	self.standardsWorksheetName = "aa_StandardsReport"
	self.standardsWorksheetKey = ""

if __name__ == "__main__":

    KOA = KAOC.KhanOAuth()
    KOA.initConnection()
    erf = open('missing_exercises.txt','w')

    GDC = GDocsClient()
    # Get the spreadsheet key
    feed = GDC.client.GetSpreadsheetsFeed()
    for f in feed.entry:
	if (f.title.text == GDC.spreadsheetName):
	    GDC.spreadsheetKey = f.id.text.split('/')[-1]

    # Get the worksheet keys
    worksheetFeed = GDC.client.GetWorksheetsFeed(GDC.spreadsheetKey)
    worksheetMap = {}
    for w in worksheetFeed.entry:
	worksheetMap[w.title.text] = w.id.text.split('/')[-1]
    GDC.studentIDWorksheetKey = worksheetMap[GDC.studentIDWorksheetName]
    GDC.khanExercisesWorksheetKey = worksheetMap[GDC.khanExercisesWorksheetName]
    GDC.standardsWorksheetKey = worksheetMap[GDC.standardsWorksheetName]
    
    listFeed = GDC.client.GetListFeed(GDC.spreadsheetKey, GDC.studentIDWorksheetKey)
    #entry = GDC.client.UpdateCell(1,1,"chrisretford", GDC.spreadsheetKey, GDC.studentIDWorksheetKey)
    #cellFeed = GDC.client.GetCellsFeed(GDC.spreadsheetKey, GDC.studentIDWorksheetKey)
    student_list = []
    
    for (i, entry) in enumerate(listFeed.entry):
	#print entry.cell.col.strip(), entry.cell.row.strip(), entry.cell.inputValue.strip()
	entry_list = entry.content.text.split(',')
	entry_map = {}
	for e in entry_list:
	    tokens = e.split(':')
	    entry_map[tokens[0].replace(' ','')] = tokens[1].strip()
	entry_map['lastname'] = entry.title.text.replace(' ','')
	student_list.append(entry_map)

    listFeed = GDC.client.GetListFeed(GDC.spreadsheetKey, GDC.khanExercisesWorksheetKey)
    exercise_map = {}
    act_map = {'ACT 13':[], 'ACT 16':[], 'ACT 20':[], 'ACT 24':[], 'ACT 28':[], 'ACT 33':[]}
    act_standards =  {'ACT 13':[], 'ACT 16':[], 'ACT 20':[], 'ACT 24':[], 'ACT 28':[], 'ACT 33':[]}
    act_col = {'ACT 13':2, 'ACT 16':8, 'ACT 20':14, 'ACT 24':20, 'ACT 28':26, 'ACT 33':32}
    act_crit = {'ACT 13':0, 'ACT 16':0, 'ACT 20':0, \
		    'ACT 24':0, 'ACT 28':0, 'ACT 33':0}
    
    for (i, entry) in enumerate(listFeed.entry):
	if i == 0: continue
	elif (entry.title.text.startswith("Row")): continue
	entry_list = entry.content.text.split(',')
	entry_map = {}
	for e in entry_list:
	    tokens = e.split(':')
	    entry_map[tokens[0].replace(' ','')] = tokens[1].strip()
	name = entry.title.text.replace(' ','_').lower()
	exercise_map[name] = entry_map
	act_cat = entry_map['actmath']
	if (act_map.has_key(act_cat)):
	    act_map[act_cat].append(name)
	    if (entry_map.has_key('critical')):
		act_crit[act_cat] += 1	
	    if (entry_map.has_key('actstandard')):
		if (entry_map['actstandard'] not in act_standards[act_cat]):
			act_standards[act_cat].append(entry_map['actstandard'])
    print act_standards
    student_row = 2
    for s in student_list:
	student_row += 1
        ws_title = s['lastname'] + ',' + s['firstname']
	if (s.has_key('act9th')): score = s['act9th']
	elif (s.has_key('act8th')): score = s['act8th']
	else: score = "unknown"
	GDC.client.UpdateCell(3, 1, score, GDC.spreadsheetKey, worksheetMap[ws_title])
    
	if (s.has_key('chriscoach')):
	    try:
		response = KOA.get_api_resource("/api/v1/user?email=" + s['khanusername'])
            except EOFError:
	       print
	    except Exception, e:
	        print "Error: %s" % e
        
	    print("---")
	    if (response.strip() != 'null'):
		jsonObject = json.loads(response)
		act_row = {'ACT 13':3, 'ACT 16':3, 'ACT 20':3, \
		    'ACT 24':3, 'ACT 28':3, 'ACT 33':3}
		act_totals = {}
		act_correct = {}
		crit_totals = {'ACT 13':0, 'ACT 16':0, 'ACT 20':0, \
		    'ACT 24':0, 'ACT 28':0, 'ACT 33':0}
		exercise_totals = {'ACT 13':0, 'ACT 16':0, 'ACT 20':0, \
		    'ACT 24':0, 'ACT 28':0, 'ACT 33':0}
		act_stan_coverage = {'ACT 13':0, 'ACT 16':0, 'ACT 20':0, \
		    'ACT 24':0, 'ACT 28':0, 'ACT 33':0}

                profList = jsonObject.get("proficient_exercises")
	        khanUser = jsonObject.get("email")
	        print ws_title
		toDoList = []

		for p in profList:
		    if (not exercise_map.has_key(p)):
			erf.write(p + '\n')
			erf.flush()
			continue
#		    print p, exercise_map[p]
		    if (not exercise_map[p].has_key('actstandard')): continue
		    if (not exercise_map[p].has_key('actmath')): continue
		    act_cat = exercise_map[p]['actmath']
		    act_standard = exercise_map[p]['actstandard']
		    if (act_cat == 'N/A' or act_standard == 'N/A'): continue
		    elif (act_cat == 'ACT Pre'): continue
		    if (exercise_map[p].has_key('critical')): 
			critical = 'Y'
			crit_totals[act_cat] += 1
		    else: critical = 'N'
		    exercise_totals[act_cat] += 1
		    
		    exerciseResponse = KOA.get_api_resource \
			("/api/v1/user/exercises/" + p + "?email=" + khanUser)
		    nextResponse = KOA.get_api_resource \
			("/api/v1/user/exercises/" + p + "/followup_exercises?email=" + khanUser)
		    logResponse = KOA.get_api_resource \
			("/api/v1/user/exercises/" + p + "/log?email=" + khanUser)

                    if (exerciseResponse.strip() != 'null'):
	                exerciseJSON = json.loads(exerciseResponse)
			nextList = json.loads(nextResponse)
			logList = json.loads(logResponse)
			total = len(logList)
			correct = 0
			if total > 20: 
			    start = total - 20
			    exerTotal = 20
			else: 
			    start = 0
			    exerTotal = total
			for i in range(start, start+exerTotal):
			    if logList[i]['correct']: correct += 1
			pctCorrect = int(float(correct)/exerTotal * 100)
			if (act_totals.has_key(act_standard)):
			    act_totals[act_standard] += exerTotal
			    act_correct[act_standard] += correct
			else:
			    act_totals[act_standard] = exerTotal
			    act_correct[act_standard] = correct
	                proficientDate = exerciseJSON.get("proficient_date")
			longStreak = str(exerciseJSON.get("longest_streak"))
			for n in nextList:
			    toDoList.append(n["exercise"])
			print p, proficientDate
			
			GDC.client.UpdateCell(act_row[act_cat], act_col[act_cat],\
			    act_standard, GDC.spreadsheetKey, worksheetMap[ws_title])
			GDC.client.UpdateCell(act_row[act_cat], act_col[act_cat] + 1,\
			    p.capitalize().replace('_',' '), GDC.spreadsheetKey, worksheetMap[ws_title])
			GDC.client.UpdateCell(act_row[act_cat], act_col[act_cat] + 2,\
			    critical, GDC.spreadsheetKey, worksheetMap[ws_title])
			GDC.client.UpdateCell(act_row[act_cat], act_col[act_cat] + 3,\
			    proficientDate.split('T')[0], GDC.spreadsheetKey, worksheetMap[ws_title])
			GDC.client.UpdateCell(act_row[act_cat], act_col[act_cat] + 4,\
			    str(pctCorrect), GDC.spreadsheetKey, worksheetMap[ws_title])
			GDC.client.UpdateCell(act_row[act_cat], act_col[act_cat] + 5,\
			    longStreak, GDC.spreadsheetKey, worksheetMap[ws_title])
			
			act_row[act_cat] += 1

		for k in act_row.keys():
		    
		    GDC.client.UpdateCell(act_row[k], act_col[k] + 1,\
		        " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		    GDC.client.UpdateCell(act_row[k], act_col[k],\
		        " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		    GDC.client.UpdateCell(act_row[k], act_col[k] + 2,\
		        " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		    GDC.client.UpdateCell(act_row[k], act_col[k] + 3,\
		        " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		    GDC.client.UpdateCell(act_row[k], act_col[k] + 4,\
		        " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		    GDC.client.UpdateCell(act_row[k], act_col[k] + 5,\
		        " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		    
		    act_row[k] += 1 
		    
		for k in act_row.keys():
		    
		    GDC.client.UpdateCell(act_row[k] + 1, act_col[k] + 1,\
		        "Recommended:", GDC.spreadsheetKey, worksheetMap[ws_title])
		    GDC.client.UpdateCell(act_row[k] + 1, act_col[k],\
		        " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		    GDC.client.UpdateCell(act_row[k] + 1, act_col[k] + 2,\
		        " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		    GDC.client.UpdateCell(act_row[k] + 1, act_col[k] + 3,\
		        " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		    GDC.client.UpdateCell(act_row[k] + 1, act_col[k] + 4,\
		        " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		    GDC.client.UpdateCell(act_row[k] + 1, act_col[k] + 5,\
		        " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		    
		    act_row[k] += 1 
		    
		for t in toDoList:
		    if t in profList: continue
		    if (not exercise_map.has_key(t)): 
			erf.write(t + '\n')
			erf.flush()
			continue
#		    print t, exercise_map[t]
		    if (not exercise_map[t].has_key('actstandard')): continue
		    if (not exercise_map[t].has_key('actmath')): continue
		    act_cat = exercise_map[t]['actmath']
		    act_standard = exercise_map[t]['actstandard']
		    if (act_cat == 'N/A' or act_standard == 'N/A'): continue
		    if (act_cat == 'ACT Pre'): continue

		    if (exercise_map[t].has_key('critical')): critical = 'Y'
		    else: critical = 'N'
		    
		    GDC.client.UpdateCell(act_row[act_cat], act_col[act_cat],\
		        act_standard, GDC.spreadsheetKey, worksheetMap[ws_title])
		    GDC.client.UpdateCell(act_row[act_cat], act_col[act_cat] + 1,\
		        t.capitalize().replace('_',' '), GDC.spreadsheetKey, worksheetMap[ws_title])
		    GDC.client.UpdateCell(act_row[act_cat], act_col[act_cat] + 2,\
		        critical, GDC.spreadsheetKey, worksheetMap[ws_title])
		    
		    act_row[act_cat] += 1
		
		for k in act_row.keys():
		    
		    GDC.client.UpdateCell(act_row[k], act_col[k] + 1,\
		        " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		    GDC.client.UpdateCell(act_row[k], act_col[k],\
		        " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		    GDC.client.UpdateCell(act_row[k], act_col[k] + 2,\
		        " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		    GDC.client.UpdateCell(act_row[k], act_col[k] + 3,\
		        " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		    GDC.client.UpdateCell(act_row[k], act_col[k] + 4,\
		        " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		    GDC.client.UpdateCell(act_row[k], act_col[k] + 5,\
		        " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		    
		    act_row[k] += 1

		for stan in act_totals.keys():
		    for cat in act_standards.keys():
			if stan in act_standards[cat]:
			    act_stan_coverage[cat] += 1

		GDC.client.UpdateCell(student_row, 1,\
		    s['lastname'], GDC.spreadsheetKey, GDC.standardsWorksheetKey)
		GDC.client.UpdateCell(student_row, 2,\
		    s['firstname'], GDC.spreadsheetKey, GDC.standardsWorksheetKey)
		if (s.has_key('act8th')):
		    GDC.client.UpdateCell(student_row, 3,\
			s['act8th'], GDC.spreadsheetKey, GDC.standardsWorksheetKey)
		if (s.has_key('act9th')):
		    GDC.client.UpdateCell(student_row, 4,\
			s['act9th'], GDC.spreadsheetKey, GDC.standardsWorksheetKey)
	    
		act_keys = ['ACT 13', 'ACT 16', 'ACT 20', 'ACT 24', 'ACT 28', 'ACT 33']
		act_dex = [7, 16, 28, 45, 67, 88]
		stan_totals = [10, 15, 19, 32, 26, 23]
		for a in range(0, 6):
		    act_key = act_keys[a]
		    start_dex = act_dex[a]
		    act_standards_list = act_standards[act_key]    
		    act_standards_list.sort()
		    GDC.client.UpdateCell(student_row, start_dex,\
			str(int(float(act_stan_coverage[act_key])/stan_totals[a]*100)), GDC.spreadsheetKey, GDC.standardsWorksheetKey)
		    GDC.client.UpdateCell(student_row, start_dex+1,\
			str(exercise_totals[act_key]), GDC.spreadsheetKey, GDC.standardsWorksheetKey)
		    GDC.client.UpdateCell(student_row, start_dex+2,\
			str(int(float(crit_totals[act_key])/act_crit[act_key]*100)), GDC.spreadsheetKey, GDC.standardsWorksheetKey)
		    for i in range(0, len(act_standards_list)):
			st = act_standards_list[i]
			if (not act_correct.has_key(st)): val = ' '
			else:
			    val = str(int(float(act_correct[st])/act_totals[st] * 100))
			GDC.client.UpdateCell(student_row, start_dex+3+i,\
			    val, GDC.spreadsheetKey, GDC.standardsWorksheetKey)
		    

