# -*- coding: utf-8 -*-
import os, time, re
import pyzbar.pyzbar as pyzbar
from PIL import Image

# define variables and functions
di, result, infoList = '/home/bowenfeng/Projects/audit-kit/VAT-invoice/images/', [], [['File Name', 'Version', 'Type', 'Code', 'Number', 'Amount', 'Date', 'Verification Code', 'Unknown']]
th1, th2, th3, th4 = 100, 125, 150, 175
imgList, imgList_gray, imgList_b, fList = [], [], [], os.listdir(di)
fList = sorted(fList)

def Execute():
    global result
    Preproc()
    Postproc()

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
    # Print results
    for obj in decodedObjects:
        print('Name: ', fList[i])
        print('Type: ', obj.type)
        print('Data: ', obj.data,'\n')
        
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

Execute()