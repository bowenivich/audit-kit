# -*- coding: utf-8 -*-
import xlrd,xlsxwriter
from os import listdir

# define variables
CellColumnPool = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
rows = []

# import information
ImportFileDirectory = str(input(r'''Please enter Import File Directory: (eg. "C:\Users\Alfred.Feng\Desktop\") '''))
ImportFileName = input('Please enter Import File Name: (eg. "Workbook1.xlsx" Enter to Select All) ')
if ImportFileName == "":
	Files = 1
else:
	Files = 0
ImportFileName = str(ImportFileName)
ImportFilePath = ImportFileDirectory + ImportFileName
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
elif len(StartCellColumn) == 2:
	EndCellColumn_N = (CellColumnPool.index(EndCellColumn[0]) + 1) * 26 + CellColumnPool.index(EndCellColumn[1]) + 1
else:
	print("Not Supported")
EndCellRow_N = input("Please enter End Cell Row: (eg. 99) ")
EndCellRow_N = int(EndCellRow_N)
ExportFilePath = input("Please enter Export File Path: (Directory and Name, Ending with .xlsx) ")

# define functions
def Selection():
	if Files == 0:
		if StartCellRow_N + 1 == EndCellRow_N:
			if StartCellColumn == EndCellColumn:
				print("Are you kidding me??? Just for fun???")
				exit()
			else:
				File_Row_Columns()
		else:
			if StartCellColumn == EndCellColumn:
				File_Rows_Column()
			else:
				File_Rows_Columns()
	else:
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

def File_Row_Columns():
	global rows
	ExcelFile = xlrd.open_workbook(ImportFilePath)
	Sheet = ExcelFile.sheet_by_name(SheetName)
	i = StartCellRow_N
	for j in range(StartCellColumn_N,EndCellColumn_N):
		rows.append(Sheet.cell_value(i,j))
	rows = [rows]

def File_Rows_Column():
	global rows
	ExcelFile = xlrd.open_workbook(ImportFilePath)
	Sheet = ExcelFile.sheet_by_name(SheetName)
	j = StartCellColumn_N
	for i in range(StartCellRow_N,EndCellRow_N):
		rows.append([Sheet.cell_value(i,j)])

def File_Rows_Columns():
	global rows
	ExcelFile = xlrd.open_workbook(ImportFilePath)
	Sheet = ExcelFile.sheet_by_name(SheetName)
	for i in range(StartCellRow_N,EndCellRow_N):
		row = []
		for j in range(StartCellColumn_N,EndCellColumn_N):
			row.append(Sheet.cell_value(i,j))
		rows.append(row)

def Files_Row_Column():
	global rows
	FileNameList = listdir(ImportFileDirectory)
	CheckWrongItem(FileNameList)
	for k in range(0,len(FileNameList)):
		ExcelFile = xlrd.open_workbook(ImportFilePath + FileNameList[k])
		Sheet = ExcelFile.sheet_by_name(SheetName)
		i = StartCellRow_N
		j = StartCellColumn_N
		rows.append([FileNameList[k],Sheet.cell_value(i,j)])

def Files_Row_Columns():
	global rows
	FileNameList = listdir(ImportFileDirectory)
	CheckWrongItem(FileNameList)
	for k in range(0,len(FileNameList)):
		ExcelFile = xlrd.open_workbook(ImportFilePath + FileNameList[k])
		Sheet = ExcelFile.sheet_by_name(SheetName)
		i = StartCellRow_N
		row = [FileNameList[k]]
		for j in range(StartCellColumn_N,EndCellColumn_N):
			row.append(Sheet.cell_value(i,j))
		rows.append(row)

def Files_Rows_Column():
	global rows
	FileNameList = listdir(ImportFileDirectory)
	CheckWrongItem(FileNameList)
	for k in range(0,len(FileNameList)):
		ExcelFile = xlrd.open_workbook(ImportFilePath + FileNameList[k])
		Sheet = ExcelFile.sheet_by_name(SheetName)
		j = StartCellColumn_N
		row = [FileNameList[k]]
		for i in range(StartCellRow_N,EndCellRow_N):
			row.append(Sheet.cell_value(i,j))
		rows.append(row)

def Files_Rows_Columns():
	global rows
	FileNameList = listdir(ImportFileDirectory)
	CheckWrongItem(FileNameList)
	for k in range(0,len(FileNameList)):
		ExcelFile = xlrd.open_workbook(ImportFilePath + FileNameList[k])
		Sheet = ExcelFile.sheet_by_name(SheetName)
		for i in range(StartCellRow_N,EndCellRow_N):
			row = [FileNameList[k][:-5]]
			for j in range(StartCellColumn_N,EndCellColumn_N):
				row.append(Sheet.cell_value(i,j))
			rows.append(row)

def CheckWrongItem(Item):
	WrongItem = []
	for p in range(0,len(Item)):
		if Item[p][-4:] == "xlsx":
			pass
		else:
			WrongItem.append(p)
	for q in range(0,len(WrongItem)):
		Item.remove(Item[q])

def ExportExcel(Breakdown):
	workbook = xlsxwriter.Workbook(ExportFilePath)
	worksheet = workbook.add_worksheet()
	for m in range(0,len(Breakdown)):
		for n in range(0,len(Breakdown[0])):
			worksheet.write(m,n,Breakdown[m][n])
	workbook.close()

# body
Selection()

# export
ExportExcel(rows)