import pandas as pa
import re


data = pa.read_excel('D:\Downloads\in1.xlsx')  #Read input File
r,c = data.shape    # obtain number of columns and rows in the data Frame
col = data.columns  # Store column headers
val = data.values   # Store values as multi dimensional lists

d = {}  # dictionary to store output columns as keys and row data
a = {}  # Dictionary to store student details other than test details(name, id and Chapter Tag)
n = 0   # variable counts columns that do not include test details

for i in range(0,c+1):      # This loop creates keys with empty values in dictionary d
    if '-' not in col[i]:
        d[col[i]]=[]
        n+=1
    else:

        d['Test_name']=[]
        for j in range(i,i+6):  # loop for column names with '-' in them
            s1=re.split('- ',col[j])[1]
            d[s1]=[]
        break

for i in range(0,r):    # Loop to traverse rows
    for x in range(0,n):    # Loop stores non-Test data(name, id and Chapter Tag)
        a[col[x]]=val[i,x]

    j=n     # j=n, is used to bring the pointer to the first test info for each row
    while j!=c:     # while loop till the last column of each row is traversed/accessed
        if val[i,j] == '-': # condition to skip empty data fields
            j+=1    # move to next column
            continue
        else:   # on encountering data fields with information
            test, s1 = re.split('- ', col[j])   # splitting column name to dicionary key(s1) and test name(test)
            test=test.strip()   # stripping blank spaces on either side
            for k in range(0,n):    # loop storing name, id and Chapter Tag for every test
                d[col[k]].append(a[col[k]])
            d['Test_name'].append(test)     # Storirng test name of each test
            comp=test   # comp variable is used to check if the column belongs to the same test
            while test == comp :    # checks for test name
                d[s1].append(val[i, j])     # appends the data to the respective key
                j+=1

                if j<c: # checks if the last column is reached
                    comp, s1 = re.split('- ', col[j])   # oba\taining test(comp) and key(s1) name for next column
                    comp=comp.strip()
                else:
                    break

            j-=1
        j+=1    # move to next column



""" creating frmae from dictionary 
sorting, renaming 
and exporting """

frame=pa.DataFrame(data=d)
frame.sort_values(['Test_name'],ascending=False).groupby('Name')
frame.rename(columns = {'id':'username'}, inplace = True)
print(frame)
frame.to_excel("D:\College\op.xlsx")
