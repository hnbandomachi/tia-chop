import os
import pandas as pd
import numpy as np

def numOfItems(data):
    for i in range(1, len(data)):
        if data[i][0] == "TỔNG CỘNG":
            return i     



def main():
    TC_data = pd.read_excel(os.getcwd() + '\TC.xlsx', header=5).values
    LF_data = pd.read_excel(os.getcwd() + '\LF.xlsx', header=5).values

    num_TC_item = numOfItems(TC_data)    
    num_LF_item = numOfItems(LF_data)    

    # Clean data
    TC_data = TC_data[0:num_TC_item-1]
    LF_data = LF_data[0:num_LF_item-1]

    # Get SKU and compare
    for i in range(num_TC_item-1):
        for j in range(num_LF_item-1):
            if TC_data[i][2] == LF_data[j][2]:
                print(TC_data[i][2])
        
    print("Hello users of Ngoc-Bui, this program is for you... enjoy it :)))")

if __name__ == "__main__":
    main()