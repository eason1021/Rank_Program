import numpy as np
import pandas as pd
import os

path = os.path.dirname(os.path.realpath(__file__))
for i in os.listdir(path):
    if i[-4:] == 'xlsx':
        data_path = path +'/' + i
df = pd.read_excel(data_path, sheet_name= None)

group_num = len(df.keys())
coulumn_name = df.get(list(df.keys())[0]).columns
group_num_list = []
switch = False

for i in df.keys():
    temp_list = df.get(i).to_numpy()
    group_num_list.append(temp_list.shape[0])

    if switch == False:
        switch = True
        original_list = temp_list
    else:
        original_list = np.row_stack((original_list, temp_list))

original_list = original_list.astype(str)
new_list = np.zeros((original_list.shape[0], original_list.shape[1])).astype(str)

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
new_df.to_excel('/workspaces/Code_Garden/cs_class/Test4.xlsx', index=None)
