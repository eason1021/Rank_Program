import numpy as np
import pandas as pd
import os
import sys


#Read Excel file in the same folder
if getattr(sys, 'frozen', False):
    path = os.path.dirname(os.path.abspath(sys.executable))
elif __file__:
    path = os.path.dirname(os.path.abspath(__file__))
print(path)
for i in os.listdir(path):
    if i[-4:] == 'xlsx':
        data_path = path +'/' + i
df = pd.read_excel(data_path, sheet_name= None)

#Check the number of datasheets & column's name
group_num = len(df.keys())
coulumn_name = df.get(list(df.keys())[0]).columns
group_num_list = []
switch = False

#Move the information from datasheets into single 2D array
#Record the number of each datasheet, enable program to accommodate the different number of groups.
for i in df.keys():
    temp_list = df.get(i).to_numpy()
    group_num_list.append(temp_list.shape[0])

    if switch == False:
        switch = True
        original_list = temp_list
    else:
        original_list = np.row_stack((original_list, temp_list))

#Create new array to do the S-shape sort
original_list = original_list.astype(str)
new_list = np.zeros((original_list.shape[0], original_list.shape[1])).astype(str)

#Record the maximum number from all datasheets
max_num = max(group_num_list)
count = 0

for i in range(max_num):
    #Even Number: From 0 to N
    if i % 2 == 0:
        for j in range(group_num):
            index = j+1
            temp_num = 0
            for k in range(index):
                if k == 0:
                    temp_num += i
                else :
                    temp_num += group_num_list[k-1]
            try:
                if group_num_list[j] <= i:
                    pass
                else:
                    new_list[count, :] = original_list[temp_num, :]
                    count += 1
            except:
                pass


    #Odd Number: From N to 0
    else:
        for j in range(group_num):
            temp = group_num - j
            temp_num = 0

            for k in range(temp):
                if k == 0:
                    temp_num += i
                else :
                    temp_num += group_num_list[k-1]
            
            try:
                if group_num_list[temp-1] <= i:
                    pass
                else:
                    new_list[count, :] = original_list[temp_num, :]
                    count += 1
            except:
                pass

new_df = pd.DataFrame(new_list, columns=coulumn_name)
output_path = path + "/" + 'New_Rank.xlsx'
new_df.to_excel(output_path, index=None)
