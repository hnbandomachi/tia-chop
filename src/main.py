from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
import os
import pandas as pd
import shutil
from openpyxl import load_workbook

def numOfItems(data):
    for i in range(1, len(data)):
        if data[i][0] == "TỔNG CỘNG":
            return i     

def forControl(A, B, row):
    for i in range(len(A)):
        if str(A[i]) == 'nan':
            A[i] = 0
        if str(B[i]) == 'nan':
            B[i] = 0

    if A[4] != B[4]:
        print("Ma don hang " + str(A[2]) + " bi KHAC SL dat don")
        ws['E'+str(row+5)].fill = PatternFill("solid", fgColor="00FF0000")
        ws['C'+str(row+5)].fill = PatternFill("solid", fgColor="00FF0000")

    if A[6] != B[6]:
        print("Ma don hang " + str(A[2]) + " bi KHAC SL NCC cancel")
        ws['G'+str(row+5)].fill = PatternFill("solid", fgColor="00FF0000")
        ws['C'+str(row+5)].fill = PatternFill("solid", fgColor="00FF0000")

    if A[7] != B[7]:
        print("Ma don hang " + str(A[2]) + " bi KHAC SL fail QC")
        ws['H'+str(row+5)].fill = PatternFill("solid", fgColor="00FF0000")
        ws['C'+str(row+5)].fill = PatternFill("solid", fgColor="00FF0000")

    if A[10] != B[10]:
        print("Ma don hang " + str(A[2]) + " bi KHAC Don gia nhap kho")
        ws['K'+str(row+5)].fill = PatternFill("solid", fgColor="00FF0000")
        ws['C'+str(row+5)].fill = PatternFill("solid", fgColor="00FF0000")

    if A[11] != B[11]:
        print("Ma don hang " + str(A[2]) + " bi KHAC Thanh tien nhap kho")
        ws['L'+str(row+5)].fill = PatternFill("solid", fgColor="00FF0000")
        ws['C'+str(row+5)].fill = PatternFill("solid", fgColor="00FF0000")

    if A[12] != B[12]:
        print("Ma don hang " + str(A[2]) + " bi KHAC VAT")
        ws['M'+str(row+5)].fill = PatternFill("solid", fgColor="00FF0000")
        ws['C'+str(row+5)].fill = PatternFill("solid", fgColor="00FF0000")

    if A[13] != B[13]:
        print("Ma don hang " + str(A[2]) + " bi KHAC Thanh tien thanh toan")
        ws['N'+str(row+5)].fill = PatternFill("solid", fgColor="00FF0000")
        ws['C'+str(row+5)].fill = PatternFill("solid", fgColor="00FF0000")

    if A[14] != B[14]:
        print("Ma don hang " + str(A[2]) + "Thanh tien can tru")
        ws['O'+str(row+5)].fill = PatternFill("solid", fgColor="00FF0000")
        ws['C'+str(row+5)].fill = PatternFill("solid", fgColor="00FF0000")



def main():
    global wb, ws

    shutil.copy2(os.getcwd() + '\TC.xlsx', os.getcwd() + '\ket_qua.xlsx')
    TC_data = pd.read_excel(os.getcwd() + '\TC.xlsx', header=5).values
    LF_data = pd.read_excel(os.getcwd() + '\LF.xlsx', header=5).values

    wb = load_workbook(os.getcwd() + '\ket_qua.xlsx')
    ws = wb['T4.22']

    num_TC_item = numOfItems(TC_data)    
    num_LF_item = numOfItems(LF_data)    

    # Clean data
    TC_data = TC_data[0:num_TC_item]
    LF_data = LF_data[0:num_LF_item]

    # Get SKU and compare
    for i in range(num_TC_item-1):
        for j in range(num_LF_item-1):
            if TC_data[i][2] == LF_data[j][2]:      # Consider updating new approximate comparation
                forControl(TC_data[i][:], LF_data[i][:], i)
    
    wb.save(os.getcwd() + '\ket_qua.xlsx')
    print("Hello users of Ngoc-Bui, this program is for you... enjoy it :)))")

if __name__ == "__main__":
    main()