import gspread
import time
from oauth2client.service_account import ServiceAccountCredentials
 
scope = ['https://spreadsheets.google.com/feeds',
		 'https://www.googleapis.com/auth/drive']

creds = ServiceAccountCredentials.from_json_keyfile_name('final-metrics.json', scope)

client = gspread.authorize(creds)

####################################################################################


file = client.open('College Advisor- Recruitment Metrics')    #EDIT HERE - CHANGE NAME IN PARANTHESIS WITH NAME OF SPREADSHEET

####################################################################################


statsSheet = file.add_worksheet(title="Final Metrics", rows="100", cols="20") #adds the Final Metrics tab to the spreadsheet
#statsSheet = file.worksheet("Final Metrics") #Final Metrics tab that already exists
valsSheet = file.get_worksheet(0)     #worksheet where data is currently listed from the job search (tab 1)

hiringList = valsSheet.col_values(6)
phoneList = valsSheet.col_values(8)
firstList = valsSheet.col_values(11)
secondList = valsSheet.col_values(14)
genderList = valsSheet.col_values(20)
raceList = valsSheet.col_values(21)
veteranList = valsSheet.col_values(22)
disabledList = valsSheet.col_values(23)
salaryList = valsSheet.col_values(25)

# Function to create and insert all labels and tabs
def formatStats():
	row1 = ["Hiring Stats", "", "", "Applicant Stats", "", "", "Ethnicity and Race", "", "", "Stages Stats", "", "", "Veteran Stats", "", "", "Disability Stats"]
	row2 = ["Source of Hire", "Count","", "Gender", "Count", "", "Ethnicity and Race", "Count", "", "Type", "Count", "", "Type", "Count", "", "Type", "Count"]
	row3 = ["Idealist", "", "", "Total Female:", "", "", "Hispanic", "", "", "Phone Screen", "", "", "Not a Veteren", "", "", "Not"]
	row4 = ["Indeed", "", "", "Total Male:", "", "", "White", "", "", "1st Round", "", "", "Disabled Veteren", "", "", "No Data"]
	row5 = ["Handshake", "", "", "Total NonBinary:", "", "", "American Indian/Alaska Native", "", "", "2nd Round", "", "", "Armed Forces Service Medal Veteren", "", "", "Yes"]
	row6 = ["LinkedIn Post", "", "", "Total Decline to Specify:", "", "", "Asian", "", "", "", "", "", "Veteren who served on active duty during a war or campaign", "", "", "Decline to Answer"]
	row7 = ["Edjoin", "", "", "No Data:", "", "", "Black", "", "", "", "", "", "Recently Separated Veteren"]
	row8 = ["Breakthrough", "", "", "Total Candidates:", "", "", "Decline to state", "", "", "", "", "", "Veteren (but not one of the four classes above)"]
	row9 = ["Glassdoor", "", "", "", "", "", "2 or more (non hispanic)", "", "", "", "", "", "Decline to state"]
	row10 = ["Referral", "", "", "", "", "", "Native Hawaiian/Other Pac Islander", "", "", "", "", "", "No data"]
	row11 = ["", "", "", "", "", "", "No Data"]
	row12 = []
	row13 = []
	row14 = ["Gender by Stages"]
	row15 = ["Gender", "Stages"]
	row16 = ["", "Phone Screen", "1st Round", "2nd Round"]
	row17 = ["Female"]
	row18 = ["Male"]
	row19 = ["NonBinary"]
	row20 = ["Decline to Specify"]
	row21 = ["No Data"]
	row22 = []
	row23 = []
	row24 = ["Ethnicity by Stages"]
	row25 = ["Ethnicity", "Stages"]
	row26 = ["", "Phone Screen", "1st Round", "2nd Round"]
	row27 = ["Hispanic"]
	row28 = ["White"]
	row29 = ["American Indian/Alaska Native"]
	row30 = ["Asian"]
	row31 = ["Black"]
	row32 = ["Decline to state"]
	row33 = ["2 or more (non hispanic)"]
	row34 = ["Native Hawaiian/Other Pac Islander"]
	row35 = ["No Data"]
	row36 = []
	row37 = []
	row38 = ["Gender and Ethnicity by Stages"]
	row39 = ["Ethnicity and Race", "Stages"]
	row40 = ["", "Phone Screen", "", "", "", "", "1st Round", "",  "", "", "", "2nd Round",  "", "", "", ""]
	row41 = ["", "Female", "Male", "NonBinary", "Decline to Answer", "No Data", "Female", "Male", "NonBinary", "Decline to Answer", "No Data", "Female", "Male", "NonBinary", "Decline to Answer", "No Data"]
	row42 = ["Hispanic"]
	row43 = ["White"]
	row44 = ["American Indian/Alaska Native"]
	row45 = ["Asian"]
	row46 = ["Black"]
	row47 = ["Decline to state"]
	row48 = ["2 or more (non hispanic)"]
	row49 = ["Native Hawaiian/Other Pac Islander"]
	row50 = ["No Data"]
	row51 = []
	row52 = []
	row53 = ["Time to Hire", "*Place val here*"]
	row54 = ["Mean Desired Salary"] 

	#Inserts all of the rows above into Final Metrics tab
	for i in range(1,55):
		if i == 40:
			time.sleep(105)
		statsSheet.insert_row((eval("row"+str(i))), i)

#Calculates mean salary of all applicants
def meanSalary():
	total = 0.0
	count = 0.0
	for element in salaryList:
		if "$" in element:
			fixed = element.replace("$", "")
			fixed = int(fixed.replace(",", ""))
			total += fixed
			count += 1
	total /= count
	statsSheet.update_acell('B54', total)

#Calculates number of applicants that applied
def numApplicants():
	statsSheet.update_acell('E8', len(valsSheet.col_values(2)) - 1)

#Counts number of each gender from all applicants
def genderStats():
	mcount = 0.0
	fcount = 0.0
	ncount = 0.0
	dcount = 0.0
	ocount = 0.0
	for element in genderList:
		if "Male" in element:
			mcount += 1
		if "Female" in element:
			fcount += 1
		if "NonBinary" in element:
			ncount += 1
		if "Decline to Specify" in element:
			dcount += 1
		if "No Data" in element:
			ocount += 1
	statsSheet.update_acell('E3', fcount)
	statsSheet.update_acell('E4', mcount)
	statsSheet.update_acell('E5', ncount)
	statsSheet.update_acell('E6', dcount)
	statsSheet.update_acell('E7', ocount)
	
#Counts number of applicants from each respective place they applied through
def hiringStats():
	idealcount = 0.0
	indeedcount = 0.0
	handcount = 0.0
	linkcount = 0.0
	edcount = 0.0
	breakcount = 0.0
	glasscount = 0.0
	refcount = 0.0
	for element in hiringList:
		if "Idealist" in element:
			idealcount += 1
		if "Indeed" in element:
			indeedcount += 1
		if "Handshake" in element:
			handcount += 1
		if "Link" in element:
			linkcount += 1
		if "Edjoin" in element:
			edcount += 1
		if "Breakthrough" in element:
			breakcount += 1
		if "Glassdoor" in element:
			glasscount += 1
		if "Referral" in element:
			refcount += 1
	statsSheet.update_acell('B3', idealcount)
	statsSheet.update_acell('B4', indeedcount)
	statsSheet.update_acell('B5', handcount)
	statsSheet.update_acell('B6', linkcount)
	statsSheet.update_acell('B7', edcount)
	statsSheet.update_acell('B8', breakcount)
	statsSheet.update_acell('B9', glasscount)
	statsSheet.update_acell('B10', refcount)

#Counts number of applicants that reached each stage of the hiring process
def stagesStats():
	pcount = 0.0
	fcount = 0.0
	scount = 0.0
	for element in phoneList:
		if element == "TRUE":
			pcount += 1
	for element in firstList:
		if element == "TRUE":
			fcount += 1
	for element in secondList:
		if element == "TRUE":
			scount += 1
	statsSheet.update_acell('K3', pcount)
	statsSheet.update_acell('K4', fcount)
	statsSheet.update_acell('K5', scount)

#Counts number of applicants a part of each respective race/ethnicity
def raceStats():
	hcount = 0.0
	wcount = 0.0
	aicount = 0.0
	acount = 0.0
	bcount = 0.0
	dcount = 0.0
	twocount = 0.0
	ncount = 0.0
	ndcount = 0.0
	for element in raceList:
		if "Hispanic" in element:
			hcount += 1
		if "White" in element:
			wcount += 1
		if "American Indian" in element:
			aicount += 1
		if "Asian" in element:
			acount += 1
		if "Black" in element:
			bcount += 1
		if "Decline" in element:
			dcount += 1
		if "Two" in element:
			twocount += 1
		if "Native Hawaiian" in element:
			ncount += 1
		if "No Data" in element:
			ndcount += 1
	statsSheet.update_acell('H3', hcount)
	statsSheet.update_acell('H4', wcount)
	statsSheet.update_acell('H5', aicount)
	statsSheet.update_acell('H6', acount)
	statsSheet.update_acell('H7', bcount)
	statsSheet.update_acell('H8', dcount)
	statsSheet.update_acell('H9', twocount)
	statsSheet.update_acell('H10', ncount)
	statsSheet.update_acell('H11', ndcount)

#Counts of number of applicants that are under a category of veteren
def veterenStats():
	ncount = 0.0
	dcount = 0.0
	acount = 0.0
	sacount = 0.0
	rcount = 0.0
	nacount = 0.0
	dscount = 0.0
	ndcount = 0.0
	for element in veteranList:
		if "Not a Veteran" in element:
			ncount += 1
		if "Disabled" in element:
			dcount += 1
		if "Armed Forces" in element:
			acount += 1
		if "who served" in element:
			sacount += 1
		if "Recently" in element:
			rcount += 1
		if "but not one" in element:
			nacount += 1
		if "Decline to state" in element:
			dscount += 1
		if "No Data" in element:
			ndcount += 1
	statsSheet.update_acell('N3', ncount)
	statsSheet.update_acell('N4', dcount)
	statsSheet.update_acell('N5', acount)
	statsSheet.update_acell('N6', sacount)
	statsSheet.update_acell('N7', rcount)
	statsSheet.update_acell('N8', nacount)
	statsSheet.update_acell('N9', dscount)
	statsSheet.update_acell('N10', ndcount)

#Counts number of applicants that have a disability
def disabledStats():
	ncount = 0.0
	ndcount = 0.0
	ycount = 0.0
	dcount = 0.0
	for element in disabledList:
		if "Not" in element:
			ncount += 1
		if "Data" in element:
			ndcount += 1
		if "Yes" in element:
			ycount += 1
		if "Decline" in element:
			dcount += 1
	statsSheet.update_acell('Q3', ncount)
	statsSheet.update_acell('Q4', ndcount)
	statsSheet.update_acell('Q5', ycount)
	statsSheet.update_acell('Q6', dcount)

#Counts number of applicants that identify as each gender with respect to interview stage
def gender_stagesStats():
	fpc = 0.0
	mpc = 0.0
	npc = 0.0
	dpc = 0.0
	ndpc = 0.0

	f1c = 0.0
	m1c = 0.0
	n1c = 0.0
	d1c = 0.0
	nd1c = 0.0

	f2c = 0.0
	m2c = 0.0
	n2c = 0.0
	d2c = 0.0
	nd2c = 0.0

	for index in range(len(genderList)):
		if index > 0:
			if "Female" in genderList[index] and phoneList[index] == "TRUE":
				fpc += 1
			if "Male" in genderList[index] and phoneList[index] == "TRUE":
				mpc += 1
			if "NonBinary" in genderList[index] and phoneList[index] == "TRUE":
				npc += 1
			if "Decline to Specify" in genderList[index] and phoneList[index] == "TRUE":
				dpc += 1
			if "No Data" in genderList[index] and phoneList[index] == "TRUE":
				ndpc += 1

			if "Female" in genderList[index] and firstList[index] == "TRUE":
				f1c += 1
			if "Male" in genderList[index] and firstList[index] == "TRUE":
				m1c += 1
			if "NonBinary" in genderList[index] and firstList[index] == "TRUE":
				n1c += 1
			if "Decline to Specify" in genderList[index] and firstList[index] == "TRUE":
				d1c += 1
			if "No Data" in genderList[index] and firstList[index] == "TRUE":
				nd1c += 1

			if "Female" in genderList[index] and secondList[index] == "TRUE":
				f2c += 1
			if "Male" in genderList[index] and secondList[index] == "TRUE":
				m2c += 1
			if "NonBinary" in genderList[index] and secondList[index] == "TRUE":
				n2c += 1
			if "Decline to Specify" in genderList[index] and secondList[index] == "TRUE":
				d2c += 1
			if "No Data" in genderList[index] and secondList[index] == "TRUE":
				nd2c += 1	

	statsSheet.update_acell('B17', fpc)
	statsSheet.update_acell('B18', mpc)
	statsSheet.update_acell('B19', npc)
	statsSheet.update_acell('B20', dpc)
	statsSheet.update_acell('B21', ndpc)

	statsSheet.update_acell('C17', f1c)
	statsSheet.update_acell('C18', m1c)
	statsSheet.update_acell('C19', n1c)
	statsSheet.update_acell('C20', d1c)
	statsSheet.update_acell('C21', nd1c)

	statsSheet.update_acell('D17', f2c)
	statsSheet.update_acell('D18', m2c)
	statsSheet.update_acell('D19', n2c)
	statsSheet.update_acell('D20', d2c)
	statsSheet.update_acell('D21', nd2c)

#Counts number of applicants that identify as each ethnicity with respect to interview stage
def ethnicity_stagesStats():
	hpc = 0.0
	wpc = 0.0
	aipc = 0.0
	apc = 0.0
	bpc = 0.0
	dpc = 0.0
	twopc = 0.0
	npc = 0.0
	ndpc = 0.0

	h1c = 0.0
	w1c = 0.0
	ai1c = 0.0
	a1c = 0.0
	b1c = 0.0
	d1c = 0.0
	two1c = 0.0
	n1c = 0.0
	nd1c = 0.0

	h2c = 0.0
	w2c = 0.0
	ai2c = 0.0
	a2c = 0.0
	b2c = 0.0
	d2c = 0.0
	two2c = 0.0
	n2c = 0.0
	nd2c = 0.0

	for index in range(len(raceList)):
		if index > 0:
			if "Hispanic" in raceList[index] and phoneList[index] == "TRUE":
				hpc+= 1
			if "White" in raceList[index] and phoneList[index] == "TRUE":
				wpc+= 1
			if "American Indian" in raceList[index] and phoneList[index] == "TRUE":
				aipc += 1
			if "Asian" in raceList[index] and phoneList[index] == "TRUE":
				apc += 1
			if "Black" in raceList[index] and phoneList[index] == "TRUE":
				bpc += 1
			if "Decline" in raceList[index] and phoneList[index] == "TRUE":
				dpc += 1
			if "Two" in raceList[index] and phoneList[index] == "TRUE":
				twopc += 1
			if "Native Hawaiian" in raceList[index] and phoneList[index] == "TRUE":
				npc += 1
			if "No Data" in raceList[index] and phoneList[index] == "TRUE":
				ndpc += 1

			if "Hispanic" in raceList[index] and firstList[index] == "TRUE":
				h1c+= 1
			if "White" in raceList[index] and firstList[index] == "TRUE":
				w1c+= 1
			if "American Indian" in raceList[index] and firstList[index] == "TRUE":
				ai1c += 1
			if "Asian" in raceList[index] and firstList[index] == "TRUE":
				a1c += 1
			if "Black" in raceList[index] and firstList[index] == "TRUE":
				b1c += 1
			if "Decline" in raceList[index] and firstList[index] == "TRUE":
				d1c += 1
			if "Two" in raceList[index] and firstList[index] == "TRUE":
				two1c += 1
			if "Native Hawaiian" in raceList[index] and firstList[index] == "TRUE":
				n1c += 1
			if "No Data" in raceList[index] and firstList[index] == "TRUE":
				nd1c += 1

			if "Hispanic" in raceList[index] and secondList[index] == "TRUE":
				h2c+= 1
			if "White" in raceList[index] and secondList[index] == "TRUE":
				w2c+= 1
			if "American Indian" in raceList[index] and secondList[index] == "TRUE":
				ai2c += 1
			if "Asian" in raceList[index] and secondList[index] == "TRUE":
				a2c += 1
			if "Black" in raceList[index] and secondList[index] == "TRUE":
				b2c += 1
			if "Decline" in raceList[index] and secondList[index] == "TRUE":
				d2c += 1
			if "Two" in raceList[index] and secondList[index] == "TRUE":
				two2c += 1
			if "Native Hawaiian" in raceList[index] and secondList[index] == "TRUE":
				n2c += 1
			if "No Data" in raceList[index] and secondList[index] == "TRUE":
				nd2c += 1

	statsSheet.update_acell('B27', hpc)
	statsSheet.update_acell('B28', wpc)
	statsSheet.update_acell('B29', aipc)
	statsSheet.update_acell('B30', apc)
	statsSheet.update_acell('B31', bpc)
	statsSheet.update_acell('B32', dpc)
	statsSheet.update_acell('B33', twopc)
	statsSheet.update_acell('B34', npc)
	statsSheet.update_acell('B35', ndpc)

	statsSheet.update_acell('C27', h1c)
	statsSheet.update_acell('C28', w1c)
	statsSheet.update_acell('C29', ai1c)
	statsSheet.update_acell('C30', a1c)
	statsSheet.update_acell('C31', b1c)
	statsSheet.update_acell('C32', d1c)
	statsSheet.update_acell('C33', two1c)
	statsSheet.update_acell('C34', n1c)
	statsSheet.update_acell('C35', nd1c)

	statsSheet.update_acell('D27', h2c)
	statsSheet.update_acell('D28', w2c)
	statsSheet.update_acell('D29', ai2c)
	statsSheet.update_acell('D30', a2c)
	statsSheet.update_acell('D31', b2c)
	statsSheet.update_acell('D32', d2c)
	statsSheet.update_acell('D33', two2c)
	statsSheet.update_acell('D34', n2c)
	statsSheet.update_acell('D35', nd2c)


#Calculates the values that correspond to the intersection of categories of ethnicity, gender, and stage
def stages_ethnicity_genderStats():
	hfpcount = 0.0
	hmpcount = 0.0
	hnpcount = 0.0
	hdpcount = 0.0
	hndpcount = 0.0
	hf1count = 0.0
	hm1count = 0.0
	hn1count = 0.0
	hd1count = 0.0
	hnd1count = 0.0
	hf2count = 0.0
	hm2count = 0.0
	hn2count = 0.0
	hd2count = 0.0
	hnd2count = 0.0

	wfpcount = 0.0
	wmpcount = 0.0
	wnpcount = 0.0
	wdpcount = 0.0
	wndpcount = 0.0
	wf1count = 0.0
	wm1count = 0.0
	wn1count = 0.0
	wd1count = 0.0
	wnd1count = 0.0
	wf2count = 0.0
	wm2count = 0.0
	wn2count = 0.0
	wn2count = 0.0
	wd2count = 0.0
	wnd2count = 0.0

	aifpcount = 0.0
	aimpcount = 0.0
	ainpcount = 0.0
	aidpcount = 0.0
	aindpcount = 0.0
	aif1count = 0.0
	aim1count = 0.0
	ain1count = 0.0
	aid1count = 0.0
	aind1count = 0.0
	aif2count = 0.0
	aim2count = 0.0
	ain2count = 0.0
	aid2count = 0.0
	aind2count = 0.0

	afpcount = 0.0
	ampcount = 0.0
	anpcount = 0.0
	adpcount = 0.0
	andpcount = 0.0
	af1count = 0.0
	am1count = 0.0
	an1count = 0.0
	ad1count = 0.0
	and1count = 0.0
	af2count = 0.0
	am2count = 0.0
	an2count = 0.0
	ad2count = 0.0
	and2count = 0.0

	bfpcount = 0.0
	bmpcount = 0.0
	bnpcount = 0.0
	bdpcount = 0.0
	bndpcount = 0.0
	bf1count = 0.0
	bm1count = 0.0
	bn1count = 0.0
	bd1count = 0.0
	bnd1count = 0.0
	bf2count = 0.0
	bm2count = 0.0
	bn2count = 0.0
	bd2count = 0.0
	bnd2count = 0.0

	dfpcount = 0.0
	dmpcount = 0.0
	dnpcount = 0.0
	ddpcount = 0.0
	dndpcount = 0.0
	df1count = 0.0
	dm1count = 0.0
	dn1count = 0.0
	dd1count = 0.0
	dnd1count = 0.0
	df2count = 0.0
	dm2count = 0.0
	dn2count = 0.0
	dd2count = 0.0
	dnd2count = 0.0

	twofpcount = 0.0
	twompcount = 0.0
	twonpcount = 0.0
	twodpcount = 0.0
	twondpcount = 0.0
	twof1count = 0.0
	twom1count = 0.0
	twon1count = 0.0
	twod1count = 0.0
	twond1count = 0.0
	twof2count = 0.0
	twom2count = 0.0
	twon2count = 0.0
	twod2count = 0.0
	twond2count = 0.0

	nfpcount = 0.0
	nmpcount = 0.0
	nnpcount = 0.0
	ndpcount = 0.0
	nndpcount = 0.0
	nf1count = 0.0
	nm1count = 0.0
	nn1count = 0.0
	nd1count = 0.0
	nnd1count = 0.0
	nf2count = 0.0
	nm2count = 0.0
	nn2count = 0.0
	nd2count = 0.0
	nnd2count = 0.0

	ndfpcount = 0.0
	ndmpcount = 0.0
	ndnpcount = 0.0
	nddpcount = 0.0
	ndndpcount = 0.0
	ndf1count = 0.0
	ndm1count = 0.0
	ndn1count = 0.0
	ndd1count = 0.0
	ndnd1count = 0.0
	ndf2count = 0.0
	ndm2count = 0.0
	ndn2count = 0.0
	ndd2count = 0.0
	ndnd2count = 0.0
	

	for index in range(len(raceList)):
		if index > 0:
			if "Hispanic" in raceList[index] and phoneList[index] == "TRUE" and "Female" in genderList[index]:
				hfpcount += 1
			if "Hispanic" in raceList[index] and phoneList[index] == "TRUE" and "Male" in genderList[index]:
				hmpcount += 1
			if "Hispanic" in raceList[index] and phoneList[index] == "TRUE" and "NonBinary" in genderList[index]:
				hnpcount += 1
			if "Hispanic" in raceList[index] and phoneList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				hdpcount += 1
			if "Hispanic" in raceList[index] and phoneList[index] == "TRUE" and "No Data" in genderList[index]:
				hndpcount += 1

			if "Hispanic" in raceList[index] and firstList[index] == "TRUE" and "Female" in genderList[index]:
				hf1count += 1
			if "Hispanic" in raceList[index] and firstList[index] == "TRUE" and "Male" in genderList[index]:
				hm1count += 1
			if "Hispanic" in raceList[index] and firstList[index] == "TRUE" and "NonBinary" in genderList[index]:
				hn1count += 1
			if "Hispanic" in raceList[index] and firstList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				hd1count += 1
			if "Hispanic" in raceList[index] and firstList[index] == "TRUE" and "No Data" in genderList[index]:
				hnd1count += 1

			if "Hispanic" in raceList[index] and secondList[index] == "TRUE" and "Female" in genderList[index]:
				hf2count += 1
			if "Hispanic" in raceList[index] and secondList[index] == "TRUE" and "Male" in genderList[index]:
				hm2count += 1
			if "Hispanic" in raceList[index] and secondList[index] == "TRUE" and "NonBinary" in genderList[index]:
				hn2count += 1
			if "Hispanic" in raceList[index] and secondList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				hd2count += 1
			if "Hispanic" in raceList[index] and secondList[index] == "TRUE" and "No Data" in genderList[index]:
				hnd2count += 1


			if "White" in raceList[index] and phoneList[index] == "TRUE" and "Female" in genderList[index]:
				wfpcount += 1
			if "White" in raceList[index] and phoneList[index] == "TRUE" and "Male" in genderList[index]:
				wmpcount += 1
			if "White" in raceList[index] and phoneList[index] == "TRUE" and "NonBinary" in genderList[index]:
				wnpcount += 1
			if "White" in raceList[index] and phoneList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				wdpcount += 1
			if "White" in raceList[index] and phoneList[index] == "TRUE" and "No Data" in genderList[index]:
				wndpcount += 1

			if "White" in raceList[index] and firstList[index] == "TRUE" and "Female" in genderList[index]:
				wf1count += 1
			if "White" in raceList[index] and firstList[index] == "TRUE" and "Male" in genderList[index]:
				wm1count += 1
			if "White" in raceList[index] and firstList[index] == "TRUE" and "NonBinary" in genderList[index]:
				wn1count += 1
			if "White" in raceList[index] and firstList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				wd1count += 1
			if "White" in raceList[index] and firstList[index] == "TRUE" and "No Data" in genderList[index]:
				wnd1count += 1
				
			if "White" in raceList[index] and secondList[index] == "TRUE" and "Female" in genderList[index]:
				wf2count += 1
			if "White" in raceList[index] and secondList[index] == "TRUE" and "Male" in genderList[index]:
				wm2count += 1
			if "White" in raceList[index] and secondList[index] == "TRUE" and "NonBinary" in genderList[index]:
				wn2count += 1
			if "White" in raceList[index] and secondList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				wd2count += 1
			if "White" in raceList[index] and secondList[index] == "TRUE" and "No Data" in genderList[index]:
				wnd2count += 1


			if "American Indian" in raceList[index] and phoneList[index] == "TRUE" and "Female" in genderList[index]:
				aifpcount += 1
			if "American Indian" in raceList[index] and phoneList[index] == "TRUE" and "Male" in genderList[index]:
				aimpcount += 1
			if "American Indian" in raceList[index] and phoneList[index] == "TRUE" and "NonBinary" in genderList[index]:
				ainpcount += 1
			if "American Indian" in raceList[index] and phoneList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				aidpcount += 1
			if "American Indian" in raceList[index] and phoneList[index] == "TRUE" and "No Data" in genderList[index]:
				aindpcount += 1

			if "American Indian" in raceList[index] and firstList[index] == "TRUE" and "Female" in genderList[index]:
				aif1count += 1
			if "American Indian" in raceList[index] and firstList[index] == "TRUE" and "Male" in genderList[index]:
				aim1count += 1
			if "American Indian" in raceList[index] and firstList[index] == "TRUE" and "NonBinary" in genderList[index]:
				ain1count += 1
			if "American Indian" in raceList[index] and firstList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				aid1count += 1
			if "American Indian" in raceList[index] and firstList[index] == "TRUE" and "No Data" in genderList[index]:
				aind1count += 1
				
			if "American Indian" in raceList[index] and secondList[index] == "TRUE" and "Female" in genderList[index]:
				aif2count += 1
			if "American Indian" in raceList[index] and secondList[index] == "TRUE" and "Male" in genderList[index]:
				aim2count += 1
			if "American Indian" in raceList[index] and secondList[index] == "TRUE" and "NonBinary" in genderList[index]:
				ain2count += 1
			if "American Indian" in raceList[index] and secondList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				aid2count += 1
			if "American Indian" in raceList[index] and secondList[index] == "TRUE" and "No Data" in genderList[index]:
				aind2count += 1


			if "Asian" in raceList[index] and phoneList[index] == "TRUE" and "Female" in genderList[index]:
				afpcount += 1
			if "Asian" in raceList[index] and phoneList[index] == "TRUE" and "Male" in genderList[index]:
				ampcount += 1
			if "Asian" in raceList[index] and phoneList[index] == "TRUE" and "NonBinary" in genderList[index]:
				anpcount += 1
			if "Asian" in raceList[index] and phoneList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				adpcount += 1
			if "Asian" in raceList[index] and phoneList[index] == "TRUE" and "No Data" in genderList[index]:
				andpcount += 1

			if "Asian" in raceList[index] and firstList[index] == "TRUE" and "Female" in genderList[index]:
				af1count += 1
			if "Asian" in raceList[index] and firstList[index] == "TRUE" and "Male" in genderList[index]:
				am1count += 1
			if "Asian" in raceList[index] and firstList[index] == "TRUE" and "NonBinary" in genderList[index]:
				an1count += 1
			if "Asian" in raceList[index] and firstList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				ad1count += 1
			if "Asian" in raceList[index] and firstList[index] == "TRUE" and "No Data" in genderList[index]:
				and1count += 1
				
			if "Asian" in raceList[index] and secondList[index] == "TRUE" and "Female" in genderList[index]:
				af2count += 1
			if "Asian" in raceList[index] and secondList[index] == "TRUE" and "Male" in genderList[index]:
				am2count += 1
			if "Asian" in raceList[index] and secondList[index] == "TRUE" and "NonBinary" in genderList[index]:
				an2count += 1
			if "Asian" in raceList[index] and secondList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				ad2count += 1
			if "Asian" in raceList[index] and secondList[index] == "TRUE" and "No Data" in genderList[index]:
				and2count += 1


			if "Black" in raceList[index] and phoneList[index] == "TRUE" and "Female" in genderList[index]:
				bfpcount += 1
			if "Black" in raceList[index] and phoneList[index] == "TRUE" and "Male" in genderList[index]:
				bmpcount += 1
			if "Black" in raceList[index] and phoneList[index] == "TRUE" and "NonBinary" in genderList[index]:
				bnpcount += 1
			if "Black" in raceList[index] and phoneList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				bdpcount += 1
			if "Black" in raceList[index] and phoneList[index] == "TRUE" and "No Data" in genderList[index]:
				bndpcount += 1

			if "Black" in raceList[index] and firstList[index] == "TRUE" and "Female" in genderList[index]:
				bf1count += 1
			if "Black" in raceList[index] and firstList[index] == "TRUE" and "Male" in genderList[index]:
				bm1count += 1
			if "Black" in raceList[index] and firstList[index] == "TRUE" and "NonBinary" in genderList[index]:
				bn1count += 1
			if "Black" in raceList[index] and firstList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				bd1count += 1
			if "Black" in raceList[index] and firstList[index] == "TRUE" and "No Data" in genderList[index]:
				bnd1count += 1
				
			if "Black" in raceList[index] and secondList[index] == "TRUE" and "Female" in genderList[index]:
				bf2count += 1
			if "Black" in raceList[index] and secondList[index] == "TRUE" and "Male" in genderList[index]:
				bm2count += 1
			if "Black" in raceList[index] and secondList[index] == "TRUE" and "NonBinary" in genderList[index]:
				bn2count += 1
			if "Black" in raceList[index] and secondList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				bd2count += 1
			if "Black" in raceList[index] and secondList[index] == "TRUE" and "No Data" in genderList[index]:
				bnd2count += 1
			

			if "Decline" in raceList[index] and phoneList[index] == "TRUE" and "Female" in genderList[index]:
				dfpcount += 1
			if "Decline" in raceList[index] and phoneList[index] == "TRUE" and "Male" in genderList[index]:
				dmpcount += 1
			if "Decline" in raceList[index] and phoneList[index] == "TRUE" and "NonBinary" in genderList[index]:
				dnpcount += 1
			if "Decline" in raceList[index] and phoneList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				ddpcount += 1
			if "Decline" in raceList[index] and phoneList[index] == "TRUE" and "No Data" in genderList[index]:
				dndpcount += 1

			if "Decline" in raceList[index] and firstList[index] == "TRUE" and "Female" in genderList[index]:
				df1count += 1
			if "Decline" in raceList[index] and firstList[index] == "TRUE" and "Male" in genderList[index]:
				dm1count += 1
			if "Decline" in raceList[index] and firstList[index] == "TRUE" and "NonBinary" in genderList[index]:
				dn1count += 1
			if "Decline" in raceList[index] and firstList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				dd1count += 1
			if "Decline" in raceList[index] and firstList[index] == "TRUE" and "No Data" in genderList[index]:
				dnd1count += 1
				
			if "Decline" in raceList[index] and secondList[index] == "TRUE" and "Female" in genderList[index]:
				df2count += 1
			if "Decline" in raceList[index] and secondList[index] == "TRUE" and "Male" in genderList[index]:
				dm2count += 1
			if "Decline" in raceList[index] and secondList[index] == "TRUE" and "NonBinary" in genderList[index]:
				dn2count += 1
			if "Decline" in raceList[index] and secondList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				dd2count += 1
			if "Decline" in raceList[index] and secondList[index] == "TRUE" and "No Data" in genderList[index]:
				dnd2count += 1

			
			if "Two" in raceList[index] and phoneList[index] == "TRUE" and "Female" in genderList[index]:
				twofpcount += 1
			if "Two" in raceList[index] and phoneList[index] == "TRUE" and "Male" in genderList[index]:
				twompcount += 1
			if "Two" in raceList[index] and phoneList[index] == "TRUE" and "NonBinary" in genderList[index]:
				twonpcount += 1
			if "Two" in raceList[index] and phoneList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				twodpcount += 1
			if "Two" in raceList[index] and phoneList[index] == "TRUE" and "No Data" in genderList[index]:
				twondpcount += 1

			if "Two" in raceList[index] and firstList[index] == "TRUE" and "Female" in genderList[index]:
				twof1count += 1
			if "Two" in raceList[index] and firstList[index] == "TRUE" and "Male" in genderList[index]:
				twom1count += 1
			if "Two" in raceList[index] and firstList[index] == "TRUE" and "NonBinary" in genderList[index]:
				twon1count += 1
			if "Two" in raceList[index] and firstList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				twod1count += 1
			if "Two" in raceList[index] and firstList[index] == "TRUE" and "No Data" in genderList[index]:
				twond1count += 1
				
			if "Two" in raceList[index] and secondList[index] == "TRUE" and "Female" in genderList[index]:
				twof2count += 1
			if "Two" in raceList[index] and secondList[index] == "TRUE" and "Male" in genderList[index]:
				twom2count += 1
			if "Two" in raceList[index] and secondList[index] == "TRUE" and "NonBinary" in genderList[index]:
				twon2count += 1
			if "Two" in raceList[index] and secondList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				twod2count += 1
			if "Two" in raceList[index] and secondList[index] == "TRUE" and "No Data" in genderList[index]:
				twond2count += 1


			if "Native Hawaiian" in raceList[index] and phoneList[index] == "TRUE" and "Female" in genderList[index]:
				nfpcount += 1
			if "Native Hawaiian" in raceList[index] and phoneList[index] == "TRUE" and "Male" in genderList[index]:
				nmpcount += 1
			if "Native Hawaiian" in raceList[index] and phoneList[index] == "TRUE" and "NonBinary" in genderList[index]:
				nnpcount += 1
			if "Native Hawaiian" in raceList[index] and phoneList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				ndpcount += 1
			if "Native Hawaiian" in raceList[index] and phoneList[index] == "TRUE" and "No Data" in genderList[index]:
				nndpcount += 1

			if "Native Hawaiian" in raceList[index] and firstList[index] == "TRUE" and "Female" in genderList[index]:
				nf1count += 1
			if "Native Hawaiian" in raceList[index] and firstList[index] == "TRUE" and "Male" in genderList[index]:
				nm1count += 1
			if "Native Hawaiian" in raceList[index] and firstList[index] == "TRUE" and "NonBinary" in genderList[index]:
				nn1count += 1
			if "Native Hawaiian" in raceList[index] and firstList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				nd1count += 1
			if "Native Hawaiian" in raceList[index] and firstList[index] == "TRUE" and "No Data" in genderList[index]:
				nnd1count += 1
				
			if "Native Hawaiian" in raceList[index] and secondList[index] == "TRUE" and "Female" in genderList[index]:
				nf2count += 1
			if "Native Hawaiian" in raceList[index] and secondList[index] == "TRUE" and "Male" in genderList[index]:
				nm2count += 1
			if "Native Hawaiian" in raceList[index] and secondList[index] == "TRUE" and "NonBinary" in genderList[index]:
				nn2count += 1
			if "Native Hawaiian" in raceList[index] and secondList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				nd2count += 1
			if "Native Hawaiian" in raceList[index] and secondList[index] == "TRUE" and "No Data" in genderList[index]:
				nnd2count += 1
			
			if "No Data" in raceList[index] and phoneList[index] == "TRUE" and "Female" in genderList[index]:
				ndfpcount += 1
			if "No Data" in raceList[index] and phoneList[index] == "TRUE" and "Male" in genderList[index]:
				ndmpcount += 1
			if "No Data" in raceList[index] and phoneList[index] == "TRUE" and "NonBinary" in genderList[index]:
				ndnpcount += 1
			if "No Data" in raceList[index] and phoneList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				nddpcount += 1
			if "No Data" in raceList[index] and phoneList[index] == "TRUE" and "No Data" in genderList[index]:
				ndndpcount += 1

			if "No Data" in raceList[index] and firstList[index] == "TRUE" and "Female" in genderList[index]:
				ndf1count += 1
			if "No Data" in raceList[index] and firstList[index] == "TRUE" and "Male" in genderList[index]:
				ndm1count += 1
			if "No Data" in raceList[index] and firstList[index] == "TRUE" and "NonBinary" in genderList[index]:
				ndn1count += 1
			if "No Data" in raceList[index] and firstList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				ndd1count += 1
			if "No Data" in raceList[index] and firstList[index] == "TRUE" and "No Data" in genderList[index]:
				ndnd1count += 1
				
			if "No Data" in raceList[index] and secondList[index] == "TRUE" and "Female" in genderList[index]:
				ndf2count += 1
			if "No Data" in raceList[index] and secondList[index] == "TRUE" and "Male" in genderList[index]:
				ndm2count += 1
			if "No Data" in raceList[index] and secondList[index] == "TRUE" and "NonBinary" in genderList[index]:
				ndn2count += 1
			if "No Data" in raceList[index] and secondList[index] == "TRUE" and "Decline to Specify" in genderList[index]:
				ndd2count += 1
			if "No Data" in raceList[index] and secondList[index] == "TRUE" and "No Data" in genderList[index]:
				ndnd2count += 1

	time.sleep(105)

	statsSheet.update_acell('B42', hfpcount)
	statsSheet.update_acell('B43', wfpcount)
	statsSheet.update_acell('B44', aifpcount)
	statsSheet.update_acell('B45', afpcount)
	statsSheet.update_acell('B46', bfpcount)
	statsSheet.update_acell('B47', dfpcount)
	statsSheet.update_acell('B48', twofpcount)
	statsSheet.update_acell('B49', nfpcount)
	statsSheet.update_acell('B50', ndfpcount)

	statsSheet.update_acell('C42', hmpcount)
	statsSheet.update_acell('C43', wmpcount)
	statsSheet.update_acell('C44', aimpcount)
	statsSheet.update_acell('C45', ampcount)
	statsSheet.update_acell('C46', bmpcount)
	statsSheet.update_acell('C47', dmpcount)
	statsSheet.update_acell('C48', twompcount)
	statsSheet.update_acell('C49', nmpcount)
	statsSheet.update_acell('C50', ndmpcount)

	statsSheet.update_acell('D42', hnpcount)
	statsSheet.update_acell('D43', wnpcount)
	statsSheet.update_acell('D44', ainpcount)
	statsSheet.update_acell('D45', anpcount)
	statsSheet.update_acell('D46', bnpcount)
	statsSheet.update_acell('D47', dnpcount)
	statsSheet.update_acell('D48', twonpcount)
	statsSheet.update_acell('D49', nnpcount)
	statsSheet.update_acell('D50', ndnpcount)

	statsSheet.update_acell('E42', hdpcount)
	statsSheet.update_acell('E43', wdpcount)
	statsSheet.update_acell('E44', aidpcount)
	statsSheet.update_acell('E45', adpcount)
	statsSheet.update_acell('E46', bdpcount)
	statsSheet.update_acell('E47', ddpcount)
	statsSheet.update_acell('E48', twodpcount)
	statsSheet.update_acell('E49', ndpcount)
	statsSheet.update_acell('E50', nddpcount)

	statsSheet.update_acell('F42', hndpcount)
	statsSheet.update_acell('F43', wndpcount)
	statsSheet.update_acell('F44', aindpcount)
	statsSheet.update_acell('F45', andpcount)
	statsSheet.update_acell('F46', bndpcount)
	statsSheet.update_acell('F47', dndpcount)
	statsSheet.update_acell('F48', twondpcount)
	statsSheet.update_acell('F49', nndpcount)
	statsSheet.update_acell('F50', ndndpcount)


	time.sleep(105)


	statsSheet.update_acell('G42', hf1count)
	statsSheet.update_acell('G43', wf1count)
	statsSheet.update_acell('G44', aif1count)
	statsSheet.update_acell('G45', af1count)
	statsSheet.update_acell('G46', bf1count)
	statsSheet.update_acell('G47', df1count)
	statsSheet.update_acell('G48', twof1count)
	statsSheet.update_acell('G49', nf1count)
	statsSheet.update_acell('G50', ndf1count)

	statsSheet.update_acell('H42', hm1count)
	statsSheet.update_acell('H43', wm1count)
	statsSheet.update_acell('H44', aim1count)
	statsSheet.update_acell('H45', am1count)
	statsSheet.update_acell('H46', bm1count)
	statsSheet.update_acell('H47', dm1count)
	statsSheet.update_acell('H48', twom1count)
	statsSheet.update_acell('H49', nm1count)
	statsSheet.update_acell('H50', ndm1count)

	statsSheet.update_acell('I42', hn1count)
	statsSheet.update_acell('I43', wn1count)
	statsSheet.update_acell('I44', ain1count)
	statsSheet.update_acell('I45', an1count)
	statsSheet.update_acell('I46', bn1count)
	statsSheet.update_acell('I47', dn1count)
	statsSheet.update_acell('I48', twon1count)
	statsSheet.update_acell('I49', nn1count)
	statsSheet.update_acell('I50', ndn1count)

	statsSheet.update_acell('J42', hd1count)
	statsSheet.update_acell('J43', wd1count)
	statsSheet.update_acell('J44', aid1count)
	statsSheet.update_acell('J45', ad1count)
	statsSheet.update_acell('J46', bd1count)
	statsSheet.update_acell('J47', dd1count)
	statsSheet.update_acell('J48', twod1count)
	statsSheet.update_acell('J49', nd1count)
	statsSheet.update_acell('J50', ndd1count)

	statsSheet.update_acell('K42', hnd1count)
	statsSheet.update_acell('K43', wnd1count)
	statsSheet.update_acell('K44', aind1count)
	statsSheet.update_acell('K45', and1count)
	statsSheet.update_acell('K46', bnd1count)
	statsSheet.update_acell('K47', dnd1count)
	statsSheet.update_acell('K48', twond1count)
	statsSheet.update_acell('K49', nnd1count)
	statsSheet.update_acell('K50', ndnd1count)




	statsSheet.update_acell('L42', hf2count)
	statsSheet.update_acell('L43', wf2count)
	statsSheet.update_acell('L44', aif2count)
	statsSheet.update_acell('L45', af2count)
	statsSheet.update_acell('L46', bf2count)
	statsSheet.update_acell('L47', df2count)
	statsSheet.update_acell('L48', twof2count)
	statsSheet.update_acell('L49', nf2count)
	statsSheet.update_acell('L50', ndf2count)

	statsSheet.update_acell('M42', hm2count)
	statsSheet.update_acell('M43', wm2count)
	statsSheet.update_acell('M44', aim2count)
	statsSheet.update_acell('M45', am2count)
	statsSheet.update_acell('M46', bm2count)
	statsSheet.update_acell('M47', dm2count)
	statsSheet.update_acell('M48', twom2count)
	statsSheet.update_acell('M49', nm2count)
	statsSheet.update_acell('M50', ndm2count)

	statsSheet.update_acell('N42', hn2count)
	statsSheet.update_acell('N43', wn2count)
	statsSheet.update_acell('N44', ain2count)
	statsSheet.update_acell('N45', an2count)
	statsSheet.update_acell('N46', bn2count)
	statsSheet.update_acell('N47', dn2count)
	statsSheet.update_acell('N48', twon2count)
	statsSheet.update_acell('N49', nn2count)
	statsSheet.update_acell('N50', ndn2count)

	statsSheet.update_acell('O42', hd2count)
	statsSheet.update_acell('O43', wd2count)
	statsSheet.update_acell('O44', aid2count)
	statsSheet.update_acell('O45', ad2count)
	statsSheet.update_acell('O46', bd2count)
	statsSheet.update_acell('O47', dd2count)
	statsSheet.update_acell('O48', twod2count)
	statsSheet.update_acell('O49', nd2count)
	statsSheet.update_acell('O50', ndd2count)

	statsSheet.update_acell('P42', hnd2count)
	statsSheet.update_acell('P43', wnd2count)
	statsSheet.update_acell('P44', aind2count)
	statsSheet.update_acell('P45', and2count)
	statsSheet.update_acell('P46', bnd2count)
	statsSheet.update_acell('P47', dnd2count)
	statsSheet.update_acell('P48', twond2count)
	statsSheet.update_acell('P49', nnd2count)
	statsSheet.update_acell('P50', ndnd2count)


def run():
	formatStats()
	hiringStats()
	genderStats()
	numApplicants()
	raceStats()
	stagesStats()
	veterenStats()
	disabledStats()
	meanSalary()
	gender_stagesStats()
	ethnicity_stagesStats()
	time.sleep(105)
	stages_ethnicity_genderStats()

##### Code runs the below two lines #####
run()
print("success")