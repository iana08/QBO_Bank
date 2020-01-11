import sys
import pandas as pd

def main():

	inputValid()

	filePath = sys.argv[1].split("/")

	fileName = filePath[len(filePath) - 1]

	path = ""

	for word in filePath:
		if word != fileName:
			path += word + '/'

	fileName = fileName.strip(".xlsx")

	print("File to read: " + sys.argv[1])

	xls = pd.ExcelFile(sys.argv[1])

	sheetNames = xls.sheet_names
	
	if (len(sheetNames) > 2):
		print("There are more than 2 sheets, only reading 2 sheets for now.")
		return

	df = goThroughList(xls, sheetNames)

	writeToExcel(df, sheetNames, path, fileName)

	return 0

def writeToExcel(df, sheetNames, path, fileName):
	writer = pd.ExcelWriter(path+fileName+'_Match_Not.xlsx', engine='xlsxwriter')

	for name in sheetNames:
		df[name].to_excel(writer, sheet_name=name)

	writer.save()

def goThroughList(xls, sheetNames):

	df = {}

	for name in sheetNames:
		df[name] = xls.parse(name)

	QB = sheetNames[0] # Where we are searching in 
	bank = sheetNames[1] # Control For We need To Search for

	bankMatch = ["Not Match" for i in range(len(df[bank]))]
	QBMatch = ["Not Match" for i in range(len(df[QB]))]

	df[QB]["Match"] = QBMatch
	df[bank]["Match"] = bankMatch

	prev_date = df[bank].iloc[0]["Date"]
	searchList = grabList(df[bank].iloc[0]["Date"], df[QB])

	for i in range(len(df[bank])):
		if (df[bank].iloc[i]["Date"] != prev_date):
			searchList = grabList(df[bank].iloc[i]["Date"], df[QB])

		if ( df[bank].iloc[i]["Payment"] != None ):
			number = searchAmount(searchList, df[bank].iloc[i]["Payment"], df[QB])
			if (number != -1):
				df[QB].iloc[number,df[QB].columns.get_loc("Match")] = "Match"
				df[bank].iloc[i, df[bank].columns.get_loc("Match")] = "Match"
		if (df[bank].iloc[i]["Deposit"] != None):
			number = searchDeposit(searchList, df[bank].iloc[i]["Deposit"], df[QB])
			if (number != -1):
				df[QB].iloc[number, df[QB].columns.get_loc("Match")] = "Match"
				df[bank].iloc[i, df[bank].columns.get_loc("Match")] = "Match"


	print("Done Now Writing Back to The File")

	return df

def inputValid():
	if len(sys.argv) <= 1:
		print("Ma you need a file name to run this code.\n")
		return 
	elif (len(sys.argv) > 2):
		print("Too many arguments you just need the file name and the code name\n")
		return

	if (not sys.argv[1].endswith(".xlsx")):
		print("The second file needs to be an excel file that ends in .xlsx .\n")
		return

def searchAmount(searchList, payment, monthArray):
	for number in searchList:
		if (monthArray.iloc[number]["Payment"] == payment and monthArray.iloc[number]["Match"] != "Match"):
			return number
	return -1

def searchDeposit(searchList, deposit, monthArray):
	for number in searchList:
		if (monthArray.iloc[number]["Deposit"] == deposit and monthArray.iloc[number]["Match"] != "Match"):
			return number
	return -1

# Returns a list that is from the monthArray of the date +/- 2 days from the date 
def grabList(date, monthArray):

	array = []
	day = 0
	month = 0
	secondDay = 0
	secondMonth = 0
	uMonth = 0
	lMonth = 0

	if ('/' in str(date)):
		day = int(date[3:5])
		month = int(date[0:2])
	elif ('-' in str(date)):
		day = int(date.day)
		month = int(date.month)

	lDay = day - 2
	lMonth = month
	if (lDay <= 0):
		if (month == 3):
			lDay = lDay + 28
		elif(month == 1 or month == 5 or month == 7 or month == 8 or month == 10 or month == 12):
			lDay = 30 + (day - 2)
		elif(month == 2 or month == 4 or month == 6 or month == 9 or month == 11):
			lDay = 31 + (day - 2)
		lMonth = month - 1

	uDay = day + 2
	uMonth = month
	if(uDay > 28 and month == 2):
		uDay = uDay - 28
		uMonth = 1 + month

	if(uDay >= 31):
		if(month == 1 or month == 3 or month == 5 or month == 7 or month == 8 or month == 10 or month == 12):
			uDay = uDay - 31
		elif(month == 4 or month == 6 or month == 9 or month == 11):
			uDay = uDay - 30
		uMonth = 1 + month


	secondDate = ""
	for i in range(len(monthArray)):
		secondDate = monthArray.iloc[i]["Date"]
		if ('/' in  str(secondDate)):
			secondDay = int(secondDate[3:5])
			secondMonth = int(secondDate[0:2])
		elif ('-' in str(secondDate)):
			secondDay = int(secondDate.day)
			secondMonth = int(secondDate.month)

		if ( ( ( ( secondDay > uDay) and ( secondMonth < uMonth) ) or ( (secondDay <= uDay) and (secondMonth <= uMonth) ) ) and ( ((secondDay >= lDay) and (secondMonth >= lMonth )) or ((lDay > secondDay) and (secondMonth > lMonth)) ) ) :
			array.append(i)

	return array

if __name__ == "__main__":
	main()