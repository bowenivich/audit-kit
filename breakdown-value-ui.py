# -*- coding: utf-8 -*-
import xlrd, xlsxwriter
from tkinter import *
from os import listdir
from multiprocessing import Process, Queue, freeze_support

# define variables and functions
CellColumnPool, finals = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', []

def Execute():
	global ImportFileDirectory, SheetName, StartCellColumn, StartCellColumn_N, StartCellRow_N, EndCellColumn, EndCellColumn_N, EndCellRow_N, ExportFilePath
	ImportFileDirectory = str(e_di.get())
	SheetName = str(e_sn.get())
	StartCellColumn = str(e_sc.get())
	if len(StartCellColumn) == 1:
		StartCellColumn_N = CellColumnPool.index(StartCellColumn)
	elif len(StartCellColumn) == 2:
		StartCellColumn_N = (CellColumnPool.index(StartCellColumn[0]) + 1) * 26 + CellColumnPool.index(StartCellColumn[1])
	else:
		print('Not Supported')
	StartCellRow_N = int(e_sr.get()) - 1
	EndCellColumn = str(e_ec.get())
	if len(EndCellColumn) == 1:
		EndCellColumn_N = CellColumnPool.index(EndCellColumn) + 1
	elif len(EndCellColumn) == 2:
		EndCellColumn_N = (CellColumnPool.index(EndCellColumn[0]) + 1) * 26 + CellColumnPool.index(EndCellColumn[1]) + 1
	else:
		print('Not Supported')
	EndCellRow_N = int(e_er.get())
	ExportFilePath = str(e_rp.get())
	try:
		Selection()
		l_msg['text'] = 'Completed'
	except:
		l_msg['text'] = 'Error'


def Selection():
	global finals
	FileNameList = listdir(ImportFileDirectory)
	FileNameList = [f for f in FileNameList if f[-4:] == 'xlsx' and f[:2] != '._']
	ProcessList = []
	q = Queue()
	if StartCellRow_N + 1 == EndCellRow_N:
		if StartCellColumn == EndCellColumn:
			for f in FileNameList:
				p = Process(target = Files_Row_Column, args = (q, f, ImportFileDirectory, SheetName, StartCellRow_N, StartCellColumn_N, ))
				ProcessList.append(p)
				p.start()
		else:
			for f in FileNameList:
				p = Process(target = Files_Row_Columns, args = (q, f, ImportFileDirectory, SheetName, StartCellRow_N, StartCellColumn_N, EndCellColumn_N, ))
				ProcessList.append(p)
				p.start()
	else:
		if StartCellColumn == EndCellColumn:
			for f in FileNameList:
				p = Process(target = Files_Rows_Column, args = (q, f, ImportFileDirectory, SheetName, StartCellRow_N, EndCellRow_N, StartCellColumn_N, ))
				ProcessList.append(p)
				p.start()
		else:
			for f in FileNameList:
				p = Process(target = Files_Rows_Columns, args = (q, f, ImportFileDirectory, SheetName, StartCellRow_N, EndCellRow_N, StartCellColumn_N, EndCellColumn_N, ))
				ProcessList.append(p)
				p.start()
	for p in ProcessList:
		finals.append(q.get())
	for p in ProcessList:
		p.join()
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
	ExcelFile = xlrd.open_workbook(ImportFileDirectory + f, encoding_override = 'utf-8')
	Sheet = ExcelFile.sheet_by_name(SheetName)
	result.append(Sheet.cell_value(i, j))
	q.put(result)

def Files_Row_Columns(q, f, ImportFileDirectory, SheetName, i, StartCellColumn_N, EndCellColumn_N):
	result = [f]
	ExcelFile = xlrd.open_workbook(ImportFileDirectory + f, encoding_override = 'utf-8')
	Sheet = ExcelFile.sheet_by_name(SheetName)
	for Column in range(StartCellColumn_N, EndCellColumn_N):
		result.append(Sheet.cell_value(i, Column))
	q.put(result)

def Files_Rows_Column(q, f, ImportFileDirectory, SheetName, StartCellRow_N, EndCellRow_N, j):
	result = [f]
	ExcelFile = xlrd.open_workbook(ImportFileDirectory + f, encoding_override = 'utf-8')
	Sheet = ExcelFile.sheet_by_name(SheetName)
	for Row in range(StartCellRow_N, EndCellRow_N):
		result.append(Sheet.cell_value(Row, j))
	q.put(result)

def Files_Rows_Columns(q, f, ImportFileDirectory, SheetName, StartCellRow_N, EndCellRow_N, StartCellColumn_N, EndCellColumn_N):
	result = []
	ExcelFile = xlrd.open_workbook(ImportFileDirectory + f, encoding_override = 'utf-8')
	Sheet = ExcelFile.sheet_by_name(SheetName)
	for Row in range(StartCellRow_N, EndCellRow_N):
		row = [f]
		for Column in range(StartCellColumn_N, EndCellColumn_N):
			try:
				row.append(Sheet.cell_value(Row, Column))
			except:
				row.append('')
		result.append(row)
	q.put(result)

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
	for l in range(0, len(Breakdown)):
		for m in range(0, len(Breakdown[l])):
			for n in range(0, len(Breakdown[l][m])):
				worksheet.write(l * len(Breakdown[l]) + m, n, Breakdown[l][m][n])
	workbook.close()

if __name__ == '__main__':
	freeze_support()
	# UI
	root = Tk()
	root.title('Breakdown by Value')
	# ImportFileDirectory
	l_di = Label(root, text=r'Directory (end with "\"): ')
	l_di.grid(row=0, sticky=W)
	e_di = Entry(root, width=50)
	e_di.grid(row=0, column=1, sticky=E)
	# SheetName
	l_sn = Label(root, text=r'Sheet Name: ')
	l_sn.grid(row=1, sticky=W)
	e_sn = Entry(root, width=50)
	e_sn.grid(row=1, column=1, sticky=E)
	# StartCellColumn
	l_sc = Label(root, text=r'Start Column (eg. A): ')
	l_sc.grid(row=2, sticky=W)
	e_sc = Entry(root, width=50)
	e_sc.grid(row=2, column=1, sticky=E)
	# StartCellRow
	l_sr = Label(root, text=r'Start Row (eg. 1): ')
	l_sr.grid(row=3, sticky=W)
	e_sr = Entry(root, width=50)
	e_sr.grid(row=3, column=1, sticky=E)
	# EndCellColumn
	l_ec = Label(root, text=r'End Column (eg. ZZ): ')
	l_ec.grid(row=4, sticky=W)
	e_ec = Entry(root, width=50)
	e_ec.grid(row=4, column=1, sticky=E)
	# EndCellRow
	l_er = Label(root, text=r'End Row (eg. 99): ')
	l_er.grid(row=5, sticky=W)
	e_er = Entry(root, width=50)
	e_er.grid(row=5, column=1, sticky=E)
	# ExportFilePath
	l_rp = Label(root, text=r'Result Path (end with ".xlsx"): ')
	l_rp.grid(row=6, sticky=W)
	e_rp = Entry(root, width=50)
	e_rp.grid(row=6, column=1, sticky=E)
	# Blank
	l_bnk = Label(root, text='')
	l_bnk.grid(row=7, columnspan=2)	
	# Execute
	b_ex = Button(root, text='Execute', command=Execute, width=25)
	b_ex.grid(row=8, columnspan=2)
	# Result
	l_msg = Label(root, text='')
	l_msg.grid(row=9, columnspan=2)	
	root.mainloop()