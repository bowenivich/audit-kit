# -*- coding: utf-8 -*-
import xlrd, xlsxwriter
from tkinter import *
from os import listdir

# define variables and functions
CellColumnPool, result = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', []

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
	global FileNameList
	FileNameList = listdir(ImportFileDirectory)
	FileNameList = [f for f in FileNameList if ((f[-4:] == 'xlsx') and (f[:2] != '._') and (f[:2] != '~$'))]
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
	ExportExcel(result)

def Files_Row_Column():
	global result
	for f in FileNameList:
		result.append([f[:-5], "='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + StartCellColumn + "$" + str(StartCellRow_N + 1)])

def Files_Row_Columns():
	global result
	for f in FileNameList:
		i, row = StartCellRow_N, [f[:-5]]
		for j in range(StartCellColumn_N, EndCellColumn_N):
			if j <= 25:
				row.append("='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j]) + "$" + str(i + 1))
			else:
				row.append("='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j//26 - 1]) + str(CellColumnPool[j%26]) + "$" + str(i+1))
		result.append(row)

def Files_Rows_Column():
	global result
	for f in FileNameList:
		j, row = StartCellColumn_N, [f[:-5]]
		for i in range(StartCellRow_N, EndCellRow_N):
			if j <= 25:
				row.append("='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j]) + "$" + str(i+1))
			else:
				row.append("='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j//26 - 1]) + str(CellColumnPool[j%26]) + "$" + str(i+1))
		result.append(row)

def Files_Rows_Columns():
	global result
	for f in FileNameList:
		for i in range(StartCellRow_N, EndCellRow_N):
			row = [f[:-5]]
			for j in range(StartCellColumn_N, EndCellColumn_N):
				if j <= 25:
					row.append("='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j]) + "$" + str(i+1))
				else:
					row.append("='" + ImportFileDirectory + "[" + f + "]" + SheetName + "'!$" + str(CellColumnPool[j//26 - 1]) + str(CellColumnPool[j%26]) + "$" + str(i+1))
			result.append(row)

def ExportExcel(Breakdown):
	workbook = xlsxwriter.Workbook(ExportFilePath)
	worksheet = workbook.add_worksheet()
	for m in range(0, len(Breakdown)):
		for n in range(0, len(Breakdown[m])):
			worksheet.write(m, n, Breakdown[m][n])
	workbook.close()

# UI
root = Tk()
root.title('Breakdown by Link')
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