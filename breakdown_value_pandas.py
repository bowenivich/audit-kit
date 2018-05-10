# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
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
StartCellColumn = str(input("Please enter Start Cell Column: (eg. A; enter blank to select all) "))
if StartCellColumn == "":
	AllColumns = 1
elif len(StartCellColumn) == 1:
	AllColumns = 0
	StartCellColumn_N = CellColumnPool.index(StartCellColumn)
	EndCellColumn = str(input("Please enter End Cell Column: (eg. E) "))
	if len(EndCellColumn) == 1:
		EndCellColumn_N = CellColumnPool.index(EndCellColumn) + 1
	elif len(EndCellColumn) == 2:
		EndCellColumn_N = (CellColumnPool.index(EndCellColumn[0]) + 1) * 26 + CellColumnPool.index(EndCellColumn[1]) + 1
	else:
		print("Not Supported! ")
elif len(StartCellColumn) == 2:
	AllColumns = 0
	StartCellColumn_N = (CellColumnPool.index(StartCellColumn[0]) + 1) * 26 + CellColumnPool.index(StartCellColumn[1])
	EndCellColumn = str(input("Please enter End Cell Column: (eg. E) "))
	if len(EndCellColumn) == 1:
		EndCellColumn_N = CellColumnPool.index(EndCellColumn) + 1
	elif len(EndCellColumn) == 2:
		EndCellColumn_N = (CellColumnPool.index(EndCellColumn[0]) + 1) * 26 + CellColumnPool.index(EndCellColumn[1]) + 1
	else:
		print("Not Supported! ")
else:
	print("Not Supported! ")
StartCellRow_N = input("Please enter Start Cell Row: (eg. 2; must begin from at least the second row) ")
StartCellRow_N = int(StartCellRow_N) - 2
EndCellRow_N = input("Please enter End Cell Row: (eg. 99) ")
EndCellRow_N = int(EndCellRow_N) - 1
ExportFilePath = input("Please enter Export File Path: (Directory and Name, Ending with .xlsx) ")

# define functions
def Selection():
	if Files == 0:
		if AllColumns == 1: 
			File_Row_AllColumns()
		elif StartCellRow_N + 1 == EndCellRow_N:
			if StartCellColumn == EndCellColumn:
				print("Are you kidding me??? Just for fun??? ")
				exit()
			else:
				File_Row_Columns()
		else:
			if StartCellColumn == EndCellColumn:
				File_Rows_Column()
			else:
				File_Rows_Columns()
	else:
		if AllColumns == 1:
			Files_Row_AllColumns()
		elif StartCellRow_N + 1 == EndCellRow_N:
			if StartCellColumn == EndCellColumn:
				Files_Row_Column()
			else:
				Files_Row_Columns()
		else:
			if StartCellColumn == EndCellColumn:
				Files_Rows_Column()
			else:
				Files_Rows_Columns()

def File_Row_AllColumns(): # checked
	global rows
	dfrm = pd.read_excel(ImportFilePath,SheetName)
	width = int(dfrm.size/len(dfrm))
	height = len(dfrm)
	if EndCellRow_N <= height:
		train = np.array(dfrm.ix[StartCellRow_N:EndCellRow_N,0:width])
	else:
		train = np.array(dfrm.ix[StartCellRow_N:height,0:width])
	row = train.tolist()
	for i in range(0,len(row)):
		rows.append(row[i])

def File_Row_Columns(): # checked
	global rows
	dfrm = pd.read_excel(ImportFilePath,SheetName)
	width = int(dfrm.size/len(dfrm))
	height = len(dfrm)
	if EndCellRow_N <= height:
		if EndCellColumn_N <= width:
			train = np.array(dfrm.ix[StartCellRow_N:EndCellRow_N,StartCellColumn_N:EndCellColumn_N])
			row = train.tolist()
		else:
			train = np.array(dfrm.ix[StartCellRow_N:EndCellRow_N,StartCellColumn_N:width])
			row = train.tolist()
			for i in range(0,len(row)):
				for j in range(0,EndCellColumn_N - width):
					row[i].append("")
		for i in range(0,len(row)):
			rows.append(row[i])
	else:
		print("The row is empty! ")
		quit()

def File_Rows_Column(): # checked
	global rows
	dfrm = pd.read_excel(ImportFilePath,SheetName)
	width = int(dfrm.size/len(dfrm))
	height = len(dfrm)
	if EndCellColumn_N <= width:
		if EndCellRow_N <= height:
			train = np.array(dfrm.ix[StartCellRow_N:EndCellRow_N,StartCellColumn_N:EndCellColumn_N])
		else:
			train = np.array(dfrm.ix[StartCellRow_N:height,StartCellColumn_N:EndCellColumn_N])
		row = train.tolist()
		for i in range(0,len(row)):
			for j in range(0,len(row[i])):
				rows.append(row[i][j])
		rows = [rows]
	else:
		print("The column is empty! ")
		quit()

def File_Rows_Columns(): # checked
	global rows
	dfrm = pd.read_excel(ImportFilePath,SheetName)
	width = int(dfrm.size/len(dfrm))
	height = len(dfrm)
	if EndCellColumn_N <= width:
		if EndCellRow_N <= height:
			train = np.array(dfrm.ix[StartCellRow_N:EndCellRow_N,StartCellColumn_N:EndCellColumn_N])
		else:
			train = np.array(dfrm.ix[StartCellRow_N:height,StartCellColumn_N:EndCellColumn_N])
		row = train.tolist()
	else:
		if EndCellRow_N <= height:
			train = np.array(dfrm.ix[StartCellRow_N:EndCellRow_N,StartCellColumn_N:width])
		else:
			train = np.array(dfrm.ix[StartCellRow_N:height,StartCellColumn_N:width])
		row = train.tolist()
		for i in range(0,len(row)):
			for j in range(0,EndCellColumn_N - width):
				row[i].append("")
	for i in range(0,len(row)):
		rows.append(row[i])

def Files_Row_AllColumns(): #checked
	global rows
	FileNameList = listdir(ImportFileDirectory)
	CheckWrongItem(FileNameList)
	for k in range(0,len(FileNameList)):
		row = [FileNameList[k]]
		dfrm = pd.read_excel(ImportFilePath + FileNameList[k],SheetName)
		width = int(dfrm.size/len(dfrm))
		height = len(dfrm)
		if EndCellRow_N <= height:
			train = np.array(dfrm.ix[StartCellRow_N:EndCellRow_N,0:width])
		else:
			train = np.array(dfrm.ix[StartCellRow_N:height,0:width])
		row = train.tolist()
		for i in range(0,len(row)):
			rows.append(row[i])

def Files_Row_Column(): # checked
	global rows
	FileNameList = listdir(ImportFileDirectory)
	CheckWrongItem(FileNameList)
	for k in range(0,len(FileNameList)):
		row = [FileNameList[k]]
		dfrm = pd.read_excel(ImportFilePath + FileNameList[k],SheetName)
		width = int(dfrm.size/len(dfrm))
		height = len(dfrm)
		if EndCellColumn_N <= width:
			if EndCellRow_N <= height:
				train = np.array(dfrm.ix[StartCellRow_N:EndCellRow_N,StartCellColumn_N:EndCellColumn_N])
				row.append(train.tolist()[0][0])
				rows.append(row)
			else:
				print("The cell is empty! ")
		else:
			print("The cell is empty! ")

def Files_Row_Columns(): # checked
	global rows
	FileNameList = listdir(ImportFileDirectory)
	CheckWrongItem(FileNameList)
	for k in range(0,len(FileNameList)):
		row = [FileNameList[k]]
		dfrm = pd.read_excel(ImportFilePath + FileNameList[k],SheetName)
		width = int(dfrm.size/len(dfrm))
		height = len(dfrm)
		if EndCellRow_N <= height:
			if EndCellColumn_N <= width:
				train = np.array(dfrm.ix[StartCellRow_N:EndCellRow_N,StartCellColumn_N:EndCellColumn_N])
				row = train.tolist()
			else:
				train = np.array(dfrm.ix[StartCellRow_N:EndCellRow_N,StartCellColumn_N:width])
				row = train.tolist()
				for i in range(0,len(row)):
					for j in range(0,EndCellColumn_N - width):
						row[i].append("")
			for i in range(0,len(row)):
				rows.append(row[i])
		else:
			print("The row in " + FileNameList[k] + " is empty! ")

def Files_Rows_Column(): # checked
	global rows
	FileNameList = listdir(ImportFileDirectory)
	CheckWrongItem(FileNameList)
	for k in range(0,len(FileNameList)):
		row = [FileNameList[k]]
		dfrm = pd.read_excel(ImportFilePath + FileNameList[k],SheetName)
		width = int(dfrm.size/len(dfrm))
		height = len(dfrm)
		if EndCellColumn_N <= width:
			if EndCellRow_N <= height:
				train = np.array(dfrm.ix[StartCellRow_N:EndCellRow_N,StartCellColumn_N:EndCellColumn_N])
			else:
				train = np.array(dfrm.ix[StartCellRow_N:height,StartCellColumn_N:EndCellColumn_N])
			row = train.tolist()
			for i in range(0,len(row)):
				for j in range(0,len(row[i])):
					rows.append(row[i][j])
			rows = [rows]
		else:
			print("The column in " + FileNameList[k] + " is empty! ")

def Files_Rows_Columns(): # checked
	global rows
	FileNameList = listdir(ImportFileDirectory)
	CheckWrongItem(FileNameList)
	for k in range(0,len(FileNameList)):
		row = [FileNameList[k]]
		dfrm = pd.read_excel(ImportFilePath + FileNameList[k],SheetName)
		width = int(dfrm.size/len(dfrm))
		height = len(dfrm)
		if EndCellColumn_N <= width:
			if EndCellRow_N <= height:
				train = np.array(dfrm.ix[StartCellRow_N:EndCellRow_N,StartCellColumn_N:EndCellColumn_N])
			else:
				train = np.array(dfrm.ix[StartCellRow_N:height,StartCellColumn_N:EndCellColumn_N])
			row = train.tolist()
		else:
			if EndCellRow_N <= height:
				train = np.array(dfrm.ix[StartCellRow_N:EndCellRow_N,StartCellColumn_N:width])
			else:
				train = np.array(dfrm.ix[StartCellRow_N:height,StartCellColumn_N:width])
			row = train.tolist()
			for i in range(0,len(row)):
				for j in range(0,EndCellColumn_N - width):
					row[i].append("")
		for i in range(0,len(row)):
			rows.append(row[i])

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
	Exp = pd.DataFrame(rows)
	writer = pd.ExcelWriter(ExportFilePath)
	Exp.to_excel(writer,"Sheet1")
	writer.save()

# body
Selection()

# export
ExportExcel(rows)