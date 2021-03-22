# Author : Yandapalli Priya Manasa
# Contact :yandapalli.manasa@ltts.com
# P.s No :99003743


#-----------------------------------------------------------------------------------------------------------------#
#-----------------------------------------------------------------------------------------------------------------#

"""
In this project the data required by the user is fetched from all available sheets in a single workbook and is written in a Master Sheet which is in the same workbook. The data is also represented as a Bar Chart and added into the same Master Sheet.

There are 10 columns in each sheet, All Sheets have 3 common data i.e. Name, PS NUmber & Email. The remaining columns are unique to each sheet.
The user provides these 3 common values to search for a Data.
Using these 3 values the record is searched in all available sheets.
This record from all the sheets is appended in a list and is Written in the Master Sheet in a single row, against the name.
The recent data from the Master Sheet is Graphically Represented.
"""
#import os
import os.path
import pandas as pd
#from openpyxl import load_workbook
#from pandas.tests.io.excel.test_openpyxl import openpyxl

df=[pd.read_excel("Sheet1.xlsx"),
	pd.read_excel("Sheet2.xlsx"),
	pd.read_excel("Sheet3.xlsx"),
	pd.read_excel("Sheet4.xlsx"),
	pd.read_excel("Sheet5.xlsx")]
# reading the data from the excel sheets
# excel sheets are present in different folders
# reading the excel sheets from different folders

emp_num = list(map(str, input("Enter Employee No's : ").split(" ")))
# map() takes two arguments.
# The first one is the method to apply,
# the second one is the data to apply it to.
# By this understanding, we can see this is doing nothing but typecasting every element of the list to an integer value.
# Since map returns the data type it was applied to, the list() method applied over map() is redundant
# split() is used to create a Python list out of a string.
# If no delimiter is given, this breaks the string by spaces.
# user gives employee number i.e p.s no as input

mega_df = []
result1 = []
result = 0
for j in emp_num:
	for i in range(39):
		if(str(df[1]["number"][i]) == j):
			# fetching different locations of folders of excel sheets into one data frame
			mega_df = [df[0].iloc[[i]],df[1].iloc[[i]],df[2].iloc[[i]],df[3].iloc[[i]],df[4].iloc[[i]]]

			# Merging columns data from multiple sheets
			result = pd.concat(mega_df, axis=1, join="inner")

			# Removing duplicate columns which are name, emp id, gmail
			result = result.loc[:, ~result.columns.duplicated()]

			# Removing any additional duplicate columns which are formed accidentally by using regular expressions
			result = result[result.columns.drop(list(result.filter(regex='Unknown')))]

			# Appending this row data into a list
			result1.append(result.iloc[[0]].values.flatten().tolist())




# Checks if there is an output.xlsx file in the directory
if(os.path.isfile("output.xlsx")):
	path = "output.xlsx"
	exist = pd.read_excel(path)
	list1 = []
	for i in range(len(exist)):
		list1.append(exist.iloc[[0]].values.flatten().tolist())
	# Adding new rows into existing data
	res = list1 + result1
	print(len(res[1]))
	res1 = pd.DataFrame(res, columns = result.columns)
	print(res1.head)
	# Updating the existing excel sheet
	res1.to_excel(path,index=False)

else:
	# Creating DataFrame from the row data list
	result2 = pd.DataFrame(result1, columns=result.columns)
	# Storing dataframe into an excel sheet
	result2.to_excel("output.xlsx", index=False)
