# -*- coding: utf-8 *-*
import gdata.spreadsheet
import gdata.spreadsheet.service
import os.path, sys
import json
import KhanOAuthConnector as KAOC

EXERCISE_FORM_WIDTH = 17
STUDENT_FORM_WIDTH = 37
STANDARDS_FORM_WIDTH = 104

def get_index(row, col, wid):
	return (int(row)-1)*wid + int(col) - 1
	
class GDocsClient():

	def __init__(self):

		gdocPath = os.path.join("private","GDocs.txt")
		file = open(gdocPath).readlines()
		self.user = file[0].split()[1].strip()
		self.password = file[1].split()[1].strip()
		self.key = file[2].split()[1].strip()
		self.client = gdata.spreadsheet.service.SpreadsheetsService()
		self.client.ClientLogin(self.user, self.password)
		self.spreadsheetName = "KhanTesting"
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

	# Pull data to fill student list
	listFeed = GDC.client.GetListFeed(GDC.spreadsheetKey, GDC.studentIDWorksheetKey)
	student_list = []
	for (i, entry) in enumerate(listFeed.entry):
		entry_list = entry.content.text.split(',')
		entry_map = {}
		for e in entry_list:
			tokens = e.split(':')
			entry_map[tokens[0].replace(' ','')] = tokens[1].strip().replace(' ','')
		entry_map['cohort'] = entry.title.text.replace(' ','')
		student_list.append(entry_map)
	# each student map now looks like (with names and scores filled)
	# {'chriscoach': 'x', 'firstname': '', 'act8th': '', 'lastname': '', 'khanusername': '', 'act9th': ''}

	# Pull all exercises names, categories, and criticalness
	listFeed = GDC.client.GetListFeed(GDC.spreadsheetKey, GDC.khanExercisesWorksheetKey)
	exercise_map = {}
	# list of exercise names belonging to each category
	act_map = {'ACT 13':[], 'ACT 16':[], 'ACT 20':[], 'ACT 24':[], 'ACT 28':[], 'ACT 33':[]}
	# list of numbered ACT standards belonging to each category
	act_standards =  {'ACT 13':[], 'ACT 16':[], 'ACT 20':[], 'ACT 24':[], 'ACT 28':[], 'ACT 33':[]}
	# column offsets into the student report
	act_col = {'ACT 13':2, 'ACT 16':8, 'ACT 20':14, 'ACT 24':20, 'ACT 28':26, 'ACT 33':32}
	# number of critical exercises per category
	act_crit = {'ACT 13':0, 'ACT 16':0, 'ACT 20':0, 'ACT 24':0, 'ACT 28':0, 'ACT 33':0}
	# high end of the score range
	act_range = {'ACT 13':15, 'ACT 16':19, 'ACT 20':23, 'ACT 24':27, 'ACT 28':32, 'ACT 33':36}
	# list of critical exercises
	crit_list = []
	for (i, entry) in enumerate(listFeed.entry):
		if (entry.title.text.startswith("Row")): continue
		if (not entry.content.text): continue
		entry_list = entry.content.text.split(',')
		entry_map = {}
		for e in entry_list:
			tokens = e.split(':')
			entry_map[tokens[0].replace(' ','')] = tokens[1].strip()
		# lower case and insert underscores for consistency w/ Khan API
		name = entry.title.text.replace(' ','_').lower()
		# location GDoc
		entry_map["index"] = i
		exercise_map[name] = entry_map
		
		act_cat = entry_map['actmath']
		if (act_cat in ('ACT Pre','N/A')): continue
		if (act_map.has_key(act_cat)):
			act_map[act_cat].append(name)
		if (entry_map.has_key('critical')):
			act_crit[act_cat] += 1
			crit_list.append(name)
		if (entry_map.has_key('actstandard')):
			if (entry_map['actstandard'] not in act_standards[act_cat]):
				act_standards[act_cat].append(entry_map['actstandard'])
				
	# get all exercises from Khan
	response = KOA.get_api_resource("/api/v1/exercises")
	json_object = json.loads(response)
	parents_map = {}
	to_visit = []
	
	for j in json_object:
		node_name = j.get("name")
		parent_list = j.get("prerequisites")
		parents_map[node_name] = parent_list
		if (not parent_list): to_visit.append(node_name)
	
	children_map = {}
	for child in parents_map.keys():
		parents = parents_map[child]
		if (parents):
			for parent in parents:
				if children_map.has_key(parent):
					children_map[parent].append(child)
				else:
					children_map[parent] = [child]
	
	visited = []
	unvisited = children_map.keys()

	while (len(to_visit) > 0):
		cur_node = to_visit.pop(0)
		visited.append(cur_node)
		if (cur_node in unvisited): 
			unvisited.remove(cur_node)
		if (children_map.has_key(cur_node)):
			for child in children_map[cur_node]:
				if (child not in to_visit):
					add_child = True
					for parent in parents_map[child]:
						if (parent not in visited):
							add_child = False
					if (add_child):
						to_visit.append(child)
	q = gdata.spreadsheet.service.CellQuery()
	q['return-empty'] = 'true'
	cells = GDC.client.GetCellsFeed(GDC.spreadsheetKey, GDC.khanExercisesWorksheetKey, query=q)
	batch_request = gdata.spreadsheet.SpreadsheetsCellsFeed()
	for v in range(len(visited)):
		print v, visited[v], exercise_map[visited[v]]
		z = get_index(exercise_map[visited[v]]["index"]+2,6,EXERCISE_FORM_WIDTH)
		cells.entry[z].cell.inputValue = str(v+1)
		batch_request.AddUpdate(cells.entry[z])
	updated = GDC.client.ExecuteBatch(batch_request, cells.GetBatchLink().href)
	
	# category names
	act_keys = ['ACT 13', 'ACT 16', 'ACT 20', 'ACT 24', 'ACT 28', 'ACT 33']
	# starting column index of each category
	act_dex = [7, 17, 31, 48, 71, 92]
	# total num of standards in each category
	stan_totals = [11, 17, 19, 33, 26, 23]

	start_student_index = 54
	# student report starts 2 rows down
	student_row = 2 + start_student_index
	for s in student_list[start_student_index:]:
		student_row += 1
		# worksheet title is lastname,firstname
		ws_title = s['lastname'] + ',' + s['firstname']
		# get the students 9th grade score preferably, 8th otherwise
		if (s.has_key('act9th')): score = s['act9th']
		elif (s.has_key('act8th')): score = s['act8th']
		else: score = "unknown"
		# initialize batch request
		q = gdata.spreadsheet.service.CellQuery()
		q['return-empty'] = 'true'
		cells = GDC.client.GetCellsFeed(GDC.spreadsheetKey, worksheetMap[ws_title], query=q)
		batch_request = gdata.spreadsheet.SpreadsheetsCellsFeed()
		# update the student report with their score
		z = get_index(1,1,STUDENT_FORM_WIDTH)
		cells.entry[z].cell.inputValue = score
		batch_request.AddUpdate(cells.entry[z])
		z = get_index(3,1,STUDENT_FORM_WIDTH)
		cells.entry[z].cell.inputValue = ' '
		batch_request.AddUpdate(cells.entry[z])
		
		#GDC.client.UpdateCell(1, 1, score, GDC.spreadsheetKey, worksheetMap[ws_title])
		#GDC.client.UpdateCell(3, 1, " ", GDC.spreadsheetKey, worksheetMap[ws_title])
		# we can only do Khan if they are my student
		if (s.has_key('chriscoach')):
			if (s['chriscoach'] != 'x'): continue
		else: continue
		# query for user info
		print("---")
		print s
		response = KOA.get_api_resource("/api/v1/user?email=" + s['khanusername'])
		if (response.strip() != 'null'):
			json_object = json.loads(response)
		else: 
			print "Couldn't get Khan data"
			continue
		# Student report starts listing exercises on row 3
		act_row = {'ACT 13':3, 'ACT 16':3, 'ACT 20':3, 'ACT 24':3, 'ACT 28':3, 'ACT 33':3}
		# how many problems did the student do per exercise
		act_totals = {}
		# how many correct problems did they get right per exercies
		act_correct = {}
		# how many critical exercises did they complete per category
		crit_totals = {'ACT 13':0, 'ACT 16':0, 'ACT 20':0, 'ACT 24':0, 'ACT 28':0, 'ACT 33':0}
		# how many exercies did they complete per category
		exercise_totals = {'ACT 13':0, 'ACT 16':0, 'ACT 20':0, 'ACT 24':0, 'ACT 28':0, 'ACT 33':0}
		# how many standards did they complete a single exersice for per category
		act_stan_coverage = {'ACT 13':0, 'ACT 16':0, 'ACT 20':0, 'ACT 24':0, 'ACT 28':0, 'ACT 33':0}
		# the difference between these is that the 'all' list has exercises that the student
		# has been given credit for without doing any exercises, e.g. Multiplaction 4 grants prof in Mult. 1
		# theyt can't be queried for streak or completion date
		all_prof_list = json_object.get("all_proficient_exercises")
		prof_list = json_object.get("proficient_exercises")
		khan_user = json_object.get("email")
		last_login = json_object.get("last_activity");
		# track the recommended exercises
		todo_list = []

		for p in all_prof_list:
			# if the student is proficient in an exercise we don't know about, write to an error file
			if (not exercise_map.has_key(p)):
				erf.write(p + '\n')
				erf.flush()
				continue
			# don't bother reporting exercises with no standard
			if (not exercise_map[p].has_key('actstandard')): continue
			if (not exercise_map[p].has_key('actmath')): continue
			# get the act category and standard for this exercise
			act_cat = exercise_map[p]['actmath']
			act_standard = exercise_map[p]['actstandard']
			# skip N/A or Pre exercies
			if (act_cat == 'N/A' or act_standard == 'N/A'): continue
			elif (act_cat == 'ACT Pre'): continue
			# set and increment criticality and exercise counts
			if (exercise_map[p].has_key('critical')): 
				critical = 'Y'
				crit_totals[act_cat] += 1
			else: critical = 'N'
			exercise_totals[act_cat] += 1
			# Get detailed information about this exercise
			exercise_response = KOA.get_api_resource("/api/v1/user/exercises/" + p + "?email=" + khan_user)
			# Get list of exercises that depend on this one
			next_response = KOA.get_api_resource("/api/v1/user/exercises/" + p + "/followup_exercises?email=" + khan_user)
			# Get log of attempted problems for this exercises
			log_response = KOA.get_api_resource("/api/v1/user/exercises/" + p + "/log?email=" + khan_user)
			if (exercise_response.strip() != 'null'):
				exercise_json = json.loads(exercise_response)
			proficient_date = exercise_json.get("proficient_date")
			if (not proficient_date): proficient_date = "2000-01-01T00:00:00Z"
			long_streak = str(exercise_json.get("longest_streak"))
			if (not long_streak): long_streak = '0'
			print p, proficient_date, long_streak
			next_list = json.loads(next_response)
			log_list = json.loads(log_response)
			# set the log range to calculate on based on last 20 or all attempted
			total = len(log_list)
			correct = 0
			if total > 20: 
				start = total - 20
				exer_total = 20
			else: 
				start = 0
				exer_total = total
			# calculate % correct
			for i in range(start, start+exer_total):
				if log_list[i]['correct']: correct += 1
			if (exer_total == 0): 
				pct_correct = 100
				exer_total = 1
				correct = 1
			else: pct_correct = int(float(correct)/exer_total * 100)
			# aggregate totals per standard
			if (act_totals.has_key(act_standard)):
				act_totals[act_standard] += exer_total
				act_correct[act_standard] += correct
			else:
				act_totals[act_standard] = exer_total
				act_correct[act_standard] = correct
			# add follow up exercises to to do list
			for n in next_list:
				if (not exercise_map.has_key(n["exercise"])):
					erf.write(n["exercise"] + '\n')
					erf.flush()
					continue
				# don't add critical exercises, they all get added later
				if (exercise_map[n["exercise"]].has_key("critical")): continue
				todo_list.append(n["exercise"])
			# populate exercise to student report
			z = get_index(act_row[act_cat], act_col[act_cat], STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = act_standard
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(act_row[act_cat], act_col[act_cat] + 1, STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = p.capitalize().replace('_', ' ')
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(act_row[act_cat], act_col[act_cat] + 2, STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = critical
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(act_row[act_cat], act_col[act_cat] + 3, STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = proficient_date.split('T')[0]
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(act_row[act_cat], act_col[act_cat] + 4, STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = str(pct_correct)
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(act_row[act_cat], act_col[act_cat] + 5, STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = long_streak
			batch_request.AddUpdate(cells.entry[z])
			act_row[act_cat] += 1
			
		# Blank row
		for k in act_row.keys():
			z = get_index(act_row[k], act_col[k], STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = ' '
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(act_row[k], act_col[k] + 1, STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = ' '
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(act_row[k], act_col[k] + 2, STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = ' '
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(act_row[k], act_col[k] + 3, STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = ' '
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(act_row[k], act_col[k] + 4, STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = ' '
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(act_row[k], act_col[k] + 5, STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = ' '
			batch_request.AddUpdate(cells.entry[z])
			act_row[k] += 1
		
		# Recommended row
		for k in act_row.keys():
			z = get_index(act_row[k], act_col[k], STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = ' '
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(act_row[k], act_col[k] + 1, STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = '** RECOMMENDED **'
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(act_row[k], act_col[k] + 2, STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = ' '
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(act_row[k], act_col[k] + 3, STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = ' '
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(act_row[k], act_col[k] + 4, STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = ' '
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(act_row[k], act_col[k] + 5, STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = ' '
			batch_request.AddUpdate(cells.entry[z])
			act_row[k] += 1
		
		# put the critical exercises in the recommended list
		todo_list.extend(crit_list)
		exam_score = int(score)
		printed_list = []
		for t in todo_list:
			# skip anything proficient
			if t in all_prof_list: continue
			# dependencies can be added multiple times, so skip any of those
			if t in printed_list: continue
			# dump unknown exercises to the error file
			if (not exercise_map.has_key(t)): 
				erf.write(t + '\n')
				erf.flush()
				continue
			# skip non-standard stuff
			if (not exercise_map[t].has_key('actstandard')): continue
			if (not exercise_map[t].has_key('actmath')): continue
			act_cat = exercise_map[t]['actmath']
			act_standard = exercise_map[t]['actstandard']
			if (act_cat == 'N/A' or act_standard == 'N/A'): continue
			if (act_cat == 'ACT Pre'): continue
			# set criticality
			if (exercise_map[t].has_key('critical')): critical = 'Y'
			else: critical = 'N'
			# don't populate non-critical below exam score
			cat_level = act_range[act_cat]
			if (critical == 'N' and exam_score > cat_level): continue
			# batch the updates to the recommended list
			z = get_index(act_row[act_cat], act_col[act_cat], STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = act_standard
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(act_row[act_cat], act_col[act_cat] + 1, STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = t.capitalize().replace('_', ' ')
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(act_row[act_cat], act_col[act_cat] + 2, STUDENT_FORM_WIDTH)
			cells.entry[z].cell.inputValue = critical
			batch_request.AddUpdate(cells.entry[z])
			printed_list.append(t)
			act_row[act_cat] += 1
		
		# Blank rows
		for i in range(0, 5):
			for k in act_row.keys():
				z = get_index(act_row[k], act_col[k], STUDENT_FORM_WIDTH)
				cells.entry[z].cell.inputValue = ' '
				batch_request.AddUpdate(cells.entry[z])
				z = get_index(act_row[k], act_col[k] + 1, STUDENT_FORM_WIDTH)
				cells.entry[z].cell.inputValue = ' '
				batch_request.AddUpdate(cells.entry[z])
				z = get_index(act_row[k], act_col[k] + 2, STUDENT_FORM_WIDTH)
				cells.entry[z].cell.inputValue = ' '
				batch_request.AddUpdate(cells.entry[z])
				z = get_index(act_row[k], act_col[k] + 3, STUDENT_FORM_WIDTH)
				cells.entry[z].cell.inputValue = ' '
				batch_request.AddUpdate(cells.entry[z])
				z = get_index(act_row[k], act_col[k] + 4, STUDENT_FORM_WIDTH)
				cells.entry[z].cell.inputValue = ' '
				batch_request.AddUpdate(cells.entry[z])
				z = get_index(act_row[k], act_col[k] + 5, STUDENT_FORM_WIDTH)
				cells.entry[z].cell.inputValue = ' '
				batch_request.AddUpdate(cells.entry[z])
				act_row[k] += 1
		# Push the batch to student report
		updated = GDC.client.ExecuteBatch(batch_request, cells.GetBatchLink().href)

		# compute the total standards covered in each category
		for stan in act_totals.keys():
			for cat in act_standards.keys():
				if stan in act_standards[cat]:
					act_stan_coverage[cat] += 1

		# initialize batch request
		q = gdata.spreadsheet.service.CellQuery()
		q['return-empty'] = 'true'
		cells = GDC.client.GetCellsFeed(GDC.spreadsheetKey, GDC.standardsWorksheetKey, query=q)
		batch_request = gdata.spreadsheet.SpreadsheetsCellsFeed()
		# fill in student name and scores
		z = get_index(student_row, 1, STANDARDS_FORM_WIDTH)
		cells.entry[z].cell.inputValue = s['lastname']
		batch_request.AddUpdate(cells.entry[z])
		z = get_index(student_row, 2, STANDARDS_FORM_WIDTH)
		cells.entry[z].cell.inputValue = s['firstname']
		batch_request.AddUpdate(cells.entry[z])
		if (s.has_key('act8th')):
			z = get_index(student_row, 3, STANDARDS_FORM_WIDTH)
			cells.entry[z].cell.inputValue = s['act8th']
			batch_request.AddUpdate(cells.entry[z])
		if (s.has_key('act9th')):
			z = get_index(student_row, 4, STANDARDS_FORM_WIDTH)
			cells.entry[z].cell.inputValue = s['act9th']
			batch_request.AddUpdate(cells.entry[z])

		# populate standards document
		for a in range(0, 6):
			# for each category, sort the standards
			act_key = act_keys[a]
			start_dex = act_dex[a]
			act_standards_list = act_standards[act_key]    
			act_standards_list.sort()
			# populate % standards covered, # exercises completed, % critical completed
			z = get_index(student_row, start_dex, STANDARDS_FORM_WIDTH)
			cells.entry[z].cell.inputValue = str(int(float(act_stan_coverage[act_key])/stan_totals[a]*100))
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(student_row, start_dex + 1, STANDARDS_FORM_WIDTH)
			cells.entry[z].cell.inputValue = str(exercise_totals[act_key])
			batch_request.AddUpdate(cells.entry[z])
			z = get_index(student_row, start_dex + 2, STANDARDS_FORM_WIDTH)
			cells.entry[z].cell.inputValue = str(int(float(crit_totals[act_key])/act_crit[act_key]*100))
			batch_request.AddUpdate(cells.entry[z])
			# populate % correct for each standard
			for i in range(0, len(act_standards_list)):
				st = act_standards_list[i]
				if (not act_correct.has_key(st)): 
					val = ' '
				else: 
					val = str(int(float(act_correct[st])/act_totals[st] * 100))
				
				z = get_index(student_row, start_dex + 3 + i, STANDARDS_FORM_WIDTH)
				cells.entry[z].cell.inputValue = val
				batch_request.AddUpdate(cells.entry[z])
		# Push the batch to standards report
		updated = GDC.client.ExecuteBatch(batch_request, cells.GetBatchLink().href)

	# initialize batch request
	q = gdata.spreadsheet.service.CellQuery()
	q['return-empty'] = 'true'
	cells = GDC.client.GetCellsFeed(GDC.spreadsheetKey, GDC.standardsWorksheetKey, query=q)
	batch_request = gdata.spreadsheet.SpreadsheetsCellsFeed()
	# fill 1st row column headers
	z = get_index(1, 1, STANDARDS_FORM_WIDTH)
	cells.entry[z].cell.inputValue = 'Student'
	batch_request.AddUpdate(cells.entry[z])
	z = get_index(1, 3, STANDARDS_FORM_WIDTH)
	cells.entry[z].cell.inputValue = 'ACT Math Score'
	batch_request.AddUpdate(cells.entry[z])
	z = get_index(1, act_dex[0], STANDARDS_FORM_WIDTH)
	cells.entry[z].cell.inputValue = 'ACT 13-15 Success Rate'
	batch_request.AddUpdate(cells.entry[z])
	z = get_index(1, act_dex[1], STANDARDS_FORM_WIDTH)
	cells.entry[z].cell.inputValue = 'ACT 16-19 Success Rate'
	batch_request.AddUpdate(cells.entry[z])
	z = get_index(1, act_dex[2], STANDARDS_FORM_WIDTH)
	cells.entry[z].cell.inputValue = 'ACT 20-23 Success Rate'
	batch_request.AddUpdate(cells.entry[z])
	z = get_index(1, act_dex[3], STANDARDS_FORM_WIDTH)
	cells.entry[z].cell.inputValue = 'ACT 24-27 Success Rate'
	batch_request.AddUpdate(cells.entry[z])
	z = get_index(1, act_dex[4], STANDARDS_FORM_WIDTH)
	cells.entry[z].cell.inputValue = 'ACT 28-32 Success Rate'
	batch_request.AddUpdate(cells.entry[z])
	z = get_index(1, act_dex[5], STANDARDS_FORM_WIDTH)
	cells.entry[z].cell.inputValue = 'ACT 33-36 Success Rate'
	batch_request.AddUpdate(cells.entry[z])
	# fill 2nd row column headers
	z = get_index(2, 1, STANDARDS_FORM_WIDTH)
	cells.entry[z].cell.inputValue = 'Last Name'
	z = get_index(2, 2, STANDARDS_FORM_WIDTH)
	cells.entry[z].cell.inputValue = 'First Name'
	z = get_index(2, 3, STANDARDS_FORM_WIDTH)
	cells.entry[z].cell.inputValue = '8th'
	z = get_index(2, 4, STANDARDS_FORM_WIDTH)
	cells.entry[z].cell.inputValue = '9th'
	z = get_index(2, 5, STANDARDS_FORM_WIDTH)
	cells.entry[z].cell.inputValue = '10th'
	z = get_index(2, 6, STANDARDS_FORM_WIDTH)
	cells.entry[z].cell.inputValue = '11th'
	batch_request.AddUpdate(cells.entry[z])
	for a in range(0, 6):
		# for each category, sort the standards
		act_key = act_keys[a]
		start_dex = act_dex[a]
		act_standards_list = act_standards[act_key]    
		act_standards_list.sort()
		# populate % standards covered, # exercises completed, % critical completed
		z = get_index(2, start_dex, STANDARDS_FORM_WIDTH)
		cells.entry[z].cell.inputValue = '% Standards Covered'
		batch_request.AddUpdate(cells.entry[z])
		z = get_index(2, start_dex + 1, STANDARDS_FORM_WIDTH)
		cells.entry[z].cell.inputValue = '# Exercises Completed'
		batch_request.AddUpdate(cells.entry[z])
		z = get_index(2, start_dex + 2, STANDARDS_FORM_WIDTH)
		cells.entry[z].cell.inputValue = '% Critical Completed'
		batch_request.AddUpdate(cells.entry[z])
		# populate correct string for each standard
		for i in range(0, len(act_standards_list)):
			z = get_index(2, start_dex + 3 + i, STANDARDS_FORM_WIDTH)
			cells.entry[z].cell.inputValue = act_standards_list[i]
			batch_request.AddUpdate(cells.entry[z])
	# Push the batch to standards report
	updated = GDC.client.ExecuteBatch(batch_request, cells.GetBatchLink().href)
