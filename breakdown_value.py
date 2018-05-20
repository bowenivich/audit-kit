# -*- coding: utf-8 -*-
import xlrd, xlsxwriter
from os import listdir
from multiprocessing import Process, Queue

# define variables
CellColumnPool, finals = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', []

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
ExportFilePath = input('Please enter Export File Path: (Directory and Name, Ending with .xlsx) ')

def Selection():
	global finals
	FileNameList = listdir(ImportFileDirectory)
	CheckWrongItem(FileNameList)
	ProcessList = []
	q = Queue()
	if StartCellRow_N + 1 == EndCellRow_N:
		if StartCellColumn == EndCellColumn:
			for f in FileNameList:
				p = Process(target=Files_Row_Column, args=(q, f, ImportFileDirectory, SheetName, StartCellRow_N, StartCellColumn_N))
				ProcessList.append(p)
				p.start()
		else:
			for f in FileNameList:
				p = Process(target=Files_Row_Columns, args=(q, f, ImportFileDirectory, SheetName, StartCellRow_N, StartCellColumn_N, EndCellColumn_N))
				ProcessList.append(p)
				p.start()
	else:
		if StartCellColumn == EndCellColumn:
			for f in FileNameList:
				p = Process(target=Files_Rows_Column, args=(q, f, ImportFileDirectory, SheetName, StartCellRow_N, EndCellRow_N, StartCellColumn_N))
				ProcessList.append(p)
				p.start()
		else:
			for f in FileNameList:
				p = Process(target=Files_Rows_Columns, args=(q, f, ImportFileDirectory, SheetName, StartCellRow_N, EndCellRow_N, StartCellColumn_N, EndCellColumn_N))
				ProcessList.append(p)
				p.start()
	for p in ProcessList:
		p.join()
	for p in ProcessList:
		finals.append(q.get())
	if StartCellRow_N + 1 == EndCellRow_N:
		if StartCellColumn == EndCellColumn:
			ExportExcel(finals)
		else:
			ExportExcel(finals)
	else:
		if StartCellColumn == EndCellColumn:
			ExportExcel(finals)
		else:
			ExportExcels(finals)

def Files_Row_Column(q, f, ImportFileDirectory, SheetName, i, j):
	result = [f]
	ExcelFile = xlrd.open_workbook(ImportFileDirectory + f)
	Sheet = ExcelFile.sheet_by_name(SheetName)
	result.append(Sheet.cell_value(i, j))
	q.put(result)

def Files_Row_Columns(q, f, ImportFileDirectory, SheetName, i, StartCellColumn_N, EndCellColumn_N):
	result = [f]
	ExcelFile = xlrd.open_workbook(ImportFileDirectory + f)
	Sheet = ExcelFile.sheet_by_name(SheetName)
	for Column in range(StartCellColumn_N, EndCellColumn_N):
		result.append(Sheet.cell_value(i, Column))
	q.put(result)

def Files_Rows_Column(q, f, ImportFileDirectory, SheetName, StartCellRow_N, EndCellRow_N, j):
	result = [f]
	ExcelFile = xlrd.open_workbook(ImportFileDirectory + f)
	Sheet = ExcelFile.sheet_by_name(SheetName)
	for Row in range(StartCellRow_N, EndCellRow_N):
		result.append(Sheet.cell_value(Row, j))
	q.put(result)

def Files_Rows_Columns(q, f, ImportFileDirectory, SheetName, StartCellRow_N, EndCellRow_N, StartCellColumn_N, EndCellColumn_N):
	result = []
	ExcelFile = xlrd.open_workbook(ImportFileDirectory + f, encoding_override='utf-8')
	Sheet = ExcelFile.sheet_by_name(SheetName)
	for Row in range(StartCellRow_N, EndCellRow_N):
		row = [f]
		for Column in range(StartCellColumn_N, EndCellColumn_N):
			row.append(Sheet.cell_value(Row, Column))
		result.append(row)
	q.put(result)

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
		for n in range(0, len(Breakdown[m])):
			worksheet.write(m, n, Breakdown[m][n])
	workbook.close()

def ExportExcels(Breakdown):
	workbook = xlsxwriter.Workbook(ExportFilePath)
	worksheet = workbook.add_worksheet()
	for l in range(0,len(Breakdown)):
		for m in range(0, len(Breakdown[l])):
			for n in range(0, len(Breakdown[l][m])):
				worksheet.write(l * len(Breakdown[l]) + m, n, Breakdown[l][m][n])
	workbook.close()

# body
Selection()