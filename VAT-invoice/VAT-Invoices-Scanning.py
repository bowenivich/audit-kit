# -*- coding: utf-8 -*-
import os, re, xlsxwriter
from pyzbar import pyzbar
from tkinter import *
from PIL import Image

# define variables and functions
result, infoList = [], [['File Name', 'Version', 'Type', 'Code', 'Number', 'Amount', 'Date', 'Verification Code', 'Unknown']]
th1, th2, th3, th4 = 100, 125, 150, 175
imgList, imgList_gray, imgList_b = [], [], []

def Execute():
    global result, fList, di, rp
    di = str(e_di.get())
    rp = str(e_rp.get())
    try:
        fList = os.listdir(di)
        fList = sorted(fList)
        Preproc()
        Postproc()
        ExportExcel(infoList)
        l_msg['text'] = 'Completed'
    except:
        l_msg['text'] = 'Error'

def Preproc():
    global fList, i, infoList
    for i in range(0, len(fList)):
        img = Image.open(di + fList[i])
        img_gray = img.convert('L')
        try:
            table = []
            for p in range(256):
                if p < th1:
                    table.append(0)
                else:
                    table.append(1)
            img_b = img_gray.point(table, '1')
            Decode(img_b)
        except:
            try:
                table = []
                for p in range(256):
                    if p < th2:
                        table.append(0)
                    else:
                        table.append(1)
                img_b = img_gray.point(table, '1')
                Decode(img_b)
            except:
                try:
                    table = []
                    for p in range(256):
                        if p < th3:
                            table.append(0)
                        else:
                            table.append(1)
                    img_b = img_gray.point(table, '1')
                    Decode(img_b)
                except:
                    try:
                        table = []
                        for p in range(256):
                            if p < th4:
                                table.append(0)
                            else:
                                table.append(1)
                        img_b = img_gray.point(table, '1')
                        Decode(img_b)
                    except:
                        result.append(['Error'])
    for i in range(0, len(result)):
        if result[i][1] == 'Error':
            infoList.append([result[i][0], 'Error'])
        else:
            info = re.findall("Decoded\(data=b'(.*?),', type", str(result[i][1]))[0].split(',')
            infoName = [result[i][0]]
            for j in range(0, len(info)):
                infoName.append(info[j])
            infoList.append(infoName)

def Decode(Image):
    # Find barcodes and QR codes
    decodedObjects = pyzbar.decode(Image)
    result.append([fList[i], decodedObjects[0]])
        
def Postproc():
    global infoList
    # Type
    for i in range(1, len(infoList)):
        if infoList[i][2] == '01':
            infoList[i][2] = '增值税专用发票'
        elif infoList[i][2] == '04':
            infoList[i][2] = '增值税普通发票'
        elif infoList[i][2] == '10':
            infoList[i][2] = '增值税电子普通发票'
        else:
            infoList[i][2] = 'Others'
    # Date
    for i in range(1, len(infoList)):
        if len(infoList[i][6]) == 8:
            infoList[i][6] = infoList[i][6][:4] + '/' + infoList[i][6][4:6] + '/' + infoList[i][6][-2:]

def ExportExcel(Info):
    workbook = xlsxwriter.Workbook(rp)
    worksheet = workbook.add_worksheet()
    for m in range(0, len(Info)):
        for n in range(0, len(Info[m]) - 1):
            worksheet.write(m, n, Info[m][n])
    workbook.close()

# UI
root = Tk()
root.title('VAT Invoices Scanning')
# Blank
l_bnk = Label(root, text='')
l_bnk.grid(row=0, columnspan=2) 
# ImportFileDirectory
l_di = Label(root, text=r'Directory (end with "\"): ')
l_di.grid(row=1, sticky=W)
e_di = Entry(root, width=50)
e_di.grid(row=1, column=1, sticky=E)
# ExportFilePath
l_rp = Label(root, text=r'Result Path (end with ".xlsx"): ')
l_rp.grid(row=2, sticky=W)
e_rp = Entry(root, width=50)
e_rp.grid(row=2, column=1, sticky=E)
# Blank
l_bnk = Label(root, text='')
l_bnk.grid(row=3, columnspan=2) 
# Execute
b_ex = Button(root, text='Execute', command=Execute, width=25)
b_ex.grid(row=4, columnspan=2)
# Result
l_msg = Label(root, text='')
l_msg.grid(row=5, columnspan=2) 
root.mainloop()