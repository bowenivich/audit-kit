# -*- coding: utf-8 -*-
import xlrd, xlsxwriter, csv
from os import listdir
from multiprocessing import Process

# define variables
CellColumnPool = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
rows = []

# import information
ImportFileDirectory = str(input(r'''Please enter Import File Directory: (eg. "C:\Users\Alfred.Feng\Desktop\") '''))
SheetName = str(input('Please enter Sheet Name: (eg. Sheet1) '))
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
TransitionDirectory = input(r'''Please enter Transition Directory: (eg. "C:\Users\Alfred.Feng\Desktop\Transition\") ''')
ExportFilePath = input('Please enter Export File Path: (Directory and Name, Ending with .xlsx) ')

# define functions
def Selection():

	FileNameList = listdir(ImportFileDirectory)
	CheckWrongItem(FileNameList)

	if StartCellRow_N + 1 == EndCellRow_N:
		if StartCellColumn == EndCellColumn:
			for f in FileNameList:
				p = Process(target=Files_Row_Column, args=(f, ImportFileDirectory, SheetName, StartCellRow_N, StartCellColumn_N))
				p.start()
		else:
			for f in FileNameList:
				p = Process(target=Files_Row_Columns, args=(f, ImportFileDirectory, SheetName, StartCellRow_N, StartCellColumn_N, EndCellColumn_N))
				p.start()
	else:
		if StartCellColumn == EndCellColumn:
			for f in FileNameList:
				p = Process(target=Files_Rows_Column, args=(f, ImportFileDirectory, SheetName, StartCellRow_N, EndCellRow_N, StartCellColumn_N))
				p.start()
		else:
			for f in FileNameList:
				p = Process(target=Files_Rows_Columns, args=(f, ImportFileDirectory, SheetName, StartCellRow_N, EndCellRow_N, StartCellColumn_N, EndCellColumn_N))
				p.start()

def Files_Row_Column(f, ImportFileDirectory, SheetName, i, j):
	result = []
	ExcelFile = xlrd.open_workbook(ImportFileDirectory + f)
	Sheet = ExcelFile.sheet_by_name(SheetName)
	result.append(Sheet.cell_value(i, j))
	with open(TransitionDirectory + f[:-5] + '.csv', 'w', newline='') as csvfile:
		spamwriter = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
		spamwriter.writerow(result)

def Files_Row_Columns(f, ImportFileDirectory, SheetName, i, StartCellColumn_N, EndCellColumn_N):
	result = []
	ExcelFile = xlrd.open_workbook(ImportFileDirectory + f)
	Sheet = ExcelFile.sheet_by_name(SheetName)
	for Column in range(StartCellColumn_N, EndCellColumn_N):
		result.append(Sheet.cell_value(i, Column))
	with open(TransitionDirectory + f[:-5] + '.csv', 'w', newline='') as csvfile:
		spamwriter = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
		spamwriter.writerow(result)

def Files_Rows_Column(f, ImportFileDirectory, SheetName, StartCellRow_N, EndCellRow_N, j):
	result = []
	ExcelFile = xlrd.open_workbook(ImportFileDirectory + f)
	Sheet = ExcelFile.sheet_by_name(SheetName)
	for Row in range(StartCellRow_N, EndCellRow_N):
		result.append(Sheet.cell_value(Row, j))
	with open(TransitionDirectory + f[:-5] + '.csv', 'w', newline='') as csvfile:
		spamwriter = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
		spamwriter.writerow(result)	

def Files_Rows_Columns(f, ImportFileDirectory, SheetName, StartCellRow_N, EndCellRow_N, StartCellColumn_N, EndCellColumn_N):
	row = []
	result = []
	ExcelFile = xlrd.open_workbook(ImportFileDirectory + f)
	Sheet = ExcelFile.sheet_by_name(SheetName)
	for Row in range(StartCellRow_N, EndCellRow_N):
		for Column in range(StartCellColumn_N, EndCellColumn_N):
			row.append(Sheet.cell_value(Row, Column))
		result.append(row)
	with open(TransitionDirectory + f[:-5] + '.csv', 'w', newline='') as csvfile:
		spamwriter = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
		for i in range(0, len(result)):
			spamwriter.writerow(result[i])
	print(f)

def CheckWrongItem(Item):
	WrongItem = []
	for p in range(0, len(Item)):
		if Item[p][-4:] == 'xlsx':
			pass
		else:
			WrongItem.append(p)
	for q in range(0, len(WrongItem)):
		Item.remove(Item[q])

def ExportExcel(Breakdown):
	workbook = xlsxwriter.Workbook(ExportFilePath)
	worksheet = workbook.add_worksheet()
	for m in range(0, len(Breakdown)):
		for n in range(0, len(Breakdown[0])):
			worksheet.write(m, n, Breakdown[m][n])
	workbook.close()

# body
Selection()

# export
# ExportExcel(rows)