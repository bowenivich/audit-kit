# -*- coding: utf-8 -*-
import xlrd, xlsxwriter
from os import listdir

# define variables
CellColumnPool, rows = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', []

# import information
ImportFileDirectory = str(input(r'''Please enter File(s) Directory and end with "\" : (eg. "C:\Users\Alfred.Feng\Engagement\") '''))
SheetName = str(input('Please enter Sheet Name: (eg. E1_应收账款) '))
StartCellColumn = str(input('Please enter Start Cell Column: (eg. A) '))
if len(StartCellColumn) == 1:
	StartCellColumn_N = CellColumnPool.index(StartCellColumn)
elif len(StartCellColumn) == 2:
	StartCellColumn_N = (CellColumnPool.index(StartCellColumn[0]) + 1) * 26 + CellColumnPool.index(StartCellColumn[1])
else:
	print('Not Supported')
StartCellRow_N = input('Please enter Start Cell Row: (eg. 1) ')
StartCellRow_N = int(StartCellRow_N) - 1
EndCellColumn = str(input('Please enter End Cell Column: (eg. E) '))
if len(EndCellColumn) == 1:
	EndCellColumn_N = CellColumnPool.index(EndCellColumn) + 1
elif len(EndCellColumn) == 2:
	EndCellColumn_N = (CellColumnPool.index(EndCellColumn[0]) + 1) * 26 + CellColumnPool.index(EndCellColumn[1]) + 1
else:
	print('Not Supported')
EndCellRow_N = input('Please enter End Cell Row: (eg. 99) ')
EndCellRow_N = int(EndCellRow_N)
ExportFilePath = input(r'''Please enter Export File Path and end with .xlsx: (eg. "C:\Users\Alfred.Feng\Engagement\result.xlsx") ''')

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
	FileNameList = [f for f in FileNameList if f[-4:] == 'xlsx' and f[:2] != '._']
	for f in FileNameList:
		i = StartCellRow_N
		j = StartCellColumn_N
		if StartCellColumn_N <= 25:
			rows.append([f[:-5], "='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j]) + "$" + str(i+1)])
		else:
			rows.append([f[:-5], "='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j//26 - 1]) + str(CellColumnPool[j%26]) + "$" + str(i+1)])

def Files_Row_Columns():
	global rows
	FileNameList = listdir(ImportFileDirectory)
	FileNameList = [f for f in FileNameList if f[-4:] == 'xlsx']
	for f in FileNameList:
		i = StartCellRow_N
		row = [f[:-5]]
		for j in range(StartCellColumn_N, EndCellColumn_N):
			if StartCellColumn_N <= 25:
				row.append("='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j]) + "$" + str(i+1))
			else:
				row.append("='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j//26 - 1]) + str(CellColumnPool[j%26]) + "$" + str(i+1))
		rows.append(row)

def Files_Rows_Column():
	global rows
	FileNameList = listdir(ImportFileDirectory)
	FileNameList = [f for f in FileNameList if f[-4:] == 'xlsx']
	for f in FileNameList:
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
	FileNameList = [f for f in FileNameList if f[-4:] == 'xlsx']
	for f in FileNameList:
		for i in range(StartCellRow_N, EndCellRow_N):
			row = [f[:-5]]
			for j in range(StartCellColumn_N, EndCellColumn_N):
				if StartCellColumn_N <= 25:
					row.append("='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j]) + "$" + str(i+1))
				else:
					row.append("='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j//26 - 1]) + str(CellColumnPool[j%26]) + "$" + str(i+1))
			rows.append(row)

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