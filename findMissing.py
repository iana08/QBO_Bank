import sys
import pandas as pd

# Global variable to store the path to the file
global excelF 
excelF = ""

# If the draw flag is passed
global draw
draw = False

def main():

	inputFlags()

	if (inputValid() == -1):
		return 0

	filePath = excelF.split("/")

	fileName = filePath[len(filePath) - 1]

	path = ""

	for word in filePath:
		if word != fileName:
			path += word + '/'

	fileName = fileName.strip(".xlsx")

	print("File to read: " + excelF)

	xls = pd.ExcelFile(excelF)

	sheetNames = xls.sheet_names
	
	if (len(sheetNames) > 2):
		print("There are more than 2 sheets, only reading 2 sheets for now.")
		return

	df = goThroughList(xls, sheetNames, draw)

	writeToExcel(df, sheetNames, path, fileName, draw)

	return 0

# Sees if the draw flag is passed and grabs the file path for the excel file
def inputFlags():
	for i in range(len(sys.argv)):
		if '-Draw' == sys.argv[i]:
			global draw 
			draw = True
		if sys.argv[i].endswith(".xlsx"):
			global excelF 
			excelF = sys.argv[i]

# Checks if we have the neccessary flags or number of flags are right
def inputValid():
	if len(sys.argv) <= 1:
		print("Ma you need a file name to run this code.\n")
		return -1
	elif (len(sys.argv) > 3):
		print("Too many arguments you just need the file name and the code name or one possible flag\n")
		return -1
	if (not excelF.endswith(".xlsx")):
		print("The second file needs to be an excel file that ends in .xlsx .\n")
		return -1

# writes to excel
def writeToExcel(df, sheetNames, path, fileName, draw):

	if draw:
		print("Writing to: ", path+fileName+'_Match_Not_Draw.xlsx')
		writer = pd.ExcelWriter(path+fileName+'_Match_Not_Draw.xlsx', engine='xlsxwriter')
	else:
		print("Writing to: ", path+fileName+'_Match_Not.xlsx')
		writer = pd.ExcelWriter(path+fileName+'_Match_Not.xlsx', engine='xlsxwriter')

	for name in sheetNames:
		df[name].to_excel(writer, sheet_name=name)

	writer.save()

def goThroughList(xls, sheetNames, draw):

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
			if not draw:	
				number = searchAmount(searchList, df[bank].iloc[i]["Payment"], df[QB])
			else:
				number = searchAmountWithDescription(searchList, df[bank].iloc[i]["Payment"], df[bank].iloc[i]["Description"], df[QB])

			if (number != -1):
				df[QB].iloc[number,df[QB].columns.get_loc("Match")] = "Match"
				df[bank].iloc[i, df[bank].columns.get_loc("Match")] = "Match"

		if (df[bank].iloc[i]["Deposit"] != None):
			if not draw:
				number = searchDeposit(searchList, df[bank].iloc[i]["Deposit"], df[QB])
			else:
				number = searchDepositWithDescription(searchList, df[bank].iloc[i]["Deposit"], df[bank].iloc[i]["Description"], df[QB])

			if (number != -1):
				df[QB].iloc[number, df[QB].columns.get_loc("Match")] = "Match"
				df[bank].iloc[i, df[bank].columns.get_loc("Match")] = "Match"

	print("Done Now Writing Back to The File")

	return df

def searchAmountWithDescription(searchList, payment, description, monthArray):
	descriptionSplit = description.split(" ")
	digits = descriptionSplit[len(descriptionSplit) - 2]
	if (digits[len(digits) - 4:] == 9680) or (digits[len(digits) - 4:] == 9568) or (digits[len(digits) - 4:] == 4581) or (digits[len(digits) - 4:] == 7039):
		for number in searchList:
			if (monthArray.iloc[number]["Payment"] == payment and monthArray.iloc[number]["Match"] != "Match"):
				if (monthArray.iloc[number]["Account"] == "Draw - TJ"):
					return number
		return -1
	else:
		return searchAmount(searchList, payment, monthArray)

def searchAmount(searchList, payment, monthArray):
	for number in searchList:
		if (monthArray.iloc[number]["Payment"] == payment and monthArray.iloc[number]["Match"] != "Match"):
			return number
	return -1

def searchDepositWithDescription(searchList, deposit, description, monthArray):
	descriptionSplit = description.split(" ")
	digits = descriptionSplit[len(descriptionSplit) - 2]
	if (digits[len(digits) - 4:] == 9680) or (digits[len(digits) - 4:] == 9568) or (digits[len(digits) - 4:] == 4581) or (digits[len(digits) - 4:] == 7039):
		for number in searchList:
			if (monthArray.iloc[number]["Deposit"] == deposit and monthArray.iloc[number]["Match"] != "Match"):
				if (monthArray.iloc[number]["Account"] == "Draw - TJ"):
					return number
		return -1
	else:
		return searchDeposit(searchList, deposit, monthArray)

def searchDeposit(searchList, deposit, monthArray):
	for number in searchList:
		if (monthArray.iloc[number]["Deposit"] == deposit and monthArray.iloc[number]["Match"] != "Match"):
			return number
	return -1

# Returns a list that is from the monthArray of the date +/- 4 days from the date 
# Will assume that the monthArray is sortted now.
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

	lDay = day - 4
	lMonth = month
	if (lDay <= 0):
		if (month == 3):
			lDay = lDay + 28
		elif(month == 1 or month == 5 or month == 7 or month == 8 or month == 10 or month == 12):
			lDay = 30 + (day - 2)
		elif(month == 2 or month == 4 or month == 6 or month == 9 or month == 11):
			lDay = 31 + (day - 2)
		lMonth = month - 1

	uDay = day + 4
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
			if monthArray.iloc[i]["Match"] != "Match":
				array.append(i)

	return array

if __name__ == "__main__":
	main()