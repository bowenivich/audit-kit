# -*- coding: utf-8 -*-
import xlrd, xlsxwriter
from os import listdir

# define variables
CellColumnPool = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
rows = []

# import information
ImportFileDirectory = str(input(r'''Please enter Import File Directory: (eg. "C:\Users\Alfred.Feng\Desktop\") '''))
SheetName = str(input("Please enter Sheet Name: (eg. Sheet1) "))
StartCellColumn = str(input("Please enter Start Cell Column: (eg. A) "))
if len(StartCellColumn) == 1:
	StartCellColumn_N = CellColumnPool.index(StartCellColumn)
elif len(StartCellColumn) == 2:
	StartCellColumn_N = (CellColumnPool.index(StartCellColumn[0]) + 1) * 26 + CellColumnPool.index(StartCellColumn[1])
else:
	print("Not Supported")
StartCellRow_N = input("Please enter Start Cell Row: (eg. 1) ")
StartCellRow_N = int(StartCellRow_N) - 1
EndCellColumn = str(input("Please enter End Cell Column: (eg. E) "))
if len(EndCellColumn) == 1:
	EndCellColumn_N = CellColumnPool.index(EndCellColumn) + 1
elif len(EndCellColumn) == 2:
	EndCellColumn_N = (CellColumnPool.index(EndCellColumn[0]) + 1) * 26 + CellColumnPool.index(EndCellColumn[1]) + 1
else:
	print("Not Supported")
EndCellRow_N = input("Please enter End Cell Row: (eg. 99) ")
EndCellRow_N = int(EndCellRow_N)
ExportFilePath = input("Please enter Export File Path: (Directory and Name, Ending with .xlsx) ")

# define functions
def Selection():
	if StartCellRow_N + 1 == EndCellRow_N:
		if StartCellColumn == EndCellColumn:
			Files_Row_Column()
		else:
			Files_Row_Columns()
	else:
		if StartCellColumn == EndCellColumn:
			Files_Rows_Column()
		else:
			Files_Rows_Columns()

def Files_Row_Column():
	global rows
	FileNameList = listdir(ImportFileDirectory)
	CheckWrongItem(FileNameList)
	for f in range(0, FileNameList):
		i = StartCellRow_N
		j = StartCellColumn_N
		if StartCellColumn_N <= 25:
			rows.append([f[:-5], "='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j]) + "$" + str(i+1)])
		else:
			rows.append([f[:-5], "='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j//26 - 1]) + str(CellColumnPool[j%26]) + "$" + str(i+1)])

def Files_Row_Columns():
	global rows
	FileNameList = listdir(ImportFileDirectory)
	CheckWrongItem(FileNameList)
	for f in range(0, FileNameList):
		i = StartCellRow_N
		row = [f:[-5]]
		for j in range(StartCellColumn_N, EndCellColumn_N):
			if StartCellColumn_N <= 25:
				row.append("='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j]) + "$" + str(i+1))
			else:
				row.append("='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j//26 - 1]) + str(CellColumnPool[j%26]) + "$" + str(i+1))
		rows.append(row)

def Files_Rows_Column():
	global rows
	FileNameList = listdir(ImportFileDirectory)
	CheckWrongItem(FileNameList)
	for k in range(0, len(FileNameList)):
		j = StartCellColumn_N
		row = [f[:-5]]
		for i in range(StartCellRow_N, EndCellRow_N):
			if StartCellColumn_N <= 25:
				row.append("='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j]) + "$" + str(i+1))
			else:
				row.append("='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j//26 - 1]) + str(CellColumnPool[j%26]) + "$" + str(i+1))
		rows.append(row)

def Files_Rows_Columns():
	global rows
	FileNameList = listdir(ImportFileDirectory)
	CheckWrongItem(FileNameList)
	for k in range(0, len(FileNameList)):
		for i in range(StartCellRow_N, EndCellRow_N):
			row = [f[:-5]]
			for j in range(StartCellColumn_N, EndCellColumn_N):
				if StartCellColumn_N <= 25:
					row.append("='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j]) + "$" + str(i+1))
				else:
					row.append("='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j//26 - 1]) + str(CellColumnPool[j%26]) + "$" + str(i+1))
			rows.append(row)

def CheckWrongItem(Item):
	WrongItem = []
	for p in range(0, len(Item)):
		if Item[p][-4:] == "xlsx":
			pass
		else:
			WrongItem.append(p)
	for q in range(0, len(WrongItem)):
		Item.remove(Item[q])

def ExportExcel(Breakdown):
	workbook = xlsxwriter.Workbook(ExportFilePath)
	worksheet = workbook.add_worksheet()
	for m in range(0, len(Breakdown)):
		for n in range(0, len(Breakdown[m])):
			worksheet.write(m, n, Breakdown[m][n])
	workbook.close()

# body
Selection()

# export
ExportExcel(rows)