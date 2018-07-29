# -*- coding: utf-8 -*-
"""
Created on Mon Dec  4 15:57:52 2017

@author: Dezter
"""
import time
import openpyxl as excel
#import numpy as np
import sys
# account time of program processing
TimeStart = time.time()  
PLM_ID = "10512007"
PLM_PW = "DES123!@#"


""" Excel file loading & Data Organize """
Location = "F:/python/ForWork/CompAprvStatus/G7C cable 0329_V0.xlsx"
wb = excel.load_workbook(Location)
ws = wb["工作表2"]
rows = ws.rows
#columns = ws.columns
contentA = []
content = []
for row in rows:
    line = [col.value for col in row if col.value is not None]
    line1 = [col.value for col in row]
    contentA.append(line)
    content.append(line1)


#Confirm each Columns meaning
ContentLen = []
for i in range(len(contentA)):
    ColLenEach = [len(j) if type(j) == str else 0 for j in contentA[i]]
    ContentLen.append(ColLenEach)
    
#find out the Column of Quanta Part Number & Manufacture Code & Manufacture Part Number
ContentLenT = list(map(list, zip(*ContentLen)))
ColLen = [max(ContentLenT[i], key=ContentLenT[i].count) for i in range(len(ContentLenT))]
Col_QPN = ColLen.index(11)
Col_MFC = ColLen.index(3)
#Col_MPN = Col_MFC+1

#find out the useful data start from which row number
QPN_Start = max(ContentLenT[Col_MFC].index(3), ContentLenT[Col_QPN].index(11))
ContentA = contentA[QPN_Start:]
Content = content[QPN_Start:]

#find out the column of Component Aprove & MD 
RowSize = [len(ContentA[i]) for i in range(len(ContentA))]
MaxCol = RowSize.index(max(RowSize))
ColSize = len(Content[MaxCol])
Col_Aprv1 = [i for i in range(len(Content[MaxCol])) if Content[MaxCol][i] == "Aprv done" or Content[MaxCol][i] == "Aprv wait" or str(Content[MaxCol][i]).startswith("A",0,1) and str(Content[MaxCol][i])[-4:-1].isdigit() == True]
Col_MD1 = [i for i in range(len(Content[MaxCol])) if Content[MaxCol][i] == "MD done" or Content[MaxCol][i] == "MD wait" or str(Content[MaxCol][i]).startswith("MD",0,2) and str(Content[MaxCol][i])[-4:-1].isdigit() == True]

if Col_Aprv1:
    Col_Aprv = Col_Aprv1[0]
elif Col_MD1:
    Col_Aprv = Col_MD1[0]-1
else :
    Col_Aprv = ColSize
    Content = [Content[i]+["0"] for i in range(len(Content))]
    
if Col_MD1:
    Col_MD = Col_MD1[0]
else :
    Col_MD = Col_Aprv+1
    Content = [Content[i]+["0"] for i in range(len(Content))]

Col_Info = [Col_QPN,Col_MFC,Col_Aprv,Col_MD]

""" Update Component Status by using Sub Program """

sys.path.append('F:/python/ForWork/CompAprvStatus')
import CompStatusCheck_SubTool as CSC
Data = CSC.CSC_SubTool(PLM_ID,PLM_PW,Col_Info,Content)

Result = content[:QPN_Start] + Data
#Write in Excel
for row in range(len(Result)):
    for col in range(len(Result[row])):
        ws.cell(row=row+1, column=col+1).value = Result[row][col]   
#Save Excel
wb.save(Location) 

# account time of program processing 
TimeEnd = time.time()
DeltaTime = TimeEnd - TimeStart
print(DeltaTime)
