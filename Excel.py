'''
read_excel: Read values from excel with the transposed shape
Input:
file - the path of the file to be read
sheet - the name of sheet to be read
row1 - the first row
row2 - the last row
col1 - the first column
col2 - the last column
sf - shedding factor
Output:
datamatrix - the numpy array with the data read from xlsx file
'''

def read_excel(file, sheet, row1, row2, col1, col2, sf):
    import xlrd
    import numpy as np
    # file = 'ACTIVSg200.xlsx'
    list_data = []
    wb = xlrd.open_workbook(filename = file)
    sheet1 = wb.sheet_by_name(sheet)
    # rows = sheet1.row_values(2)
    for j in range(col1 - 1, col2) :
        data = []
        for i in range(row1 - 1, row2) :
            if sheet1.cell(i,j).ctype == 2 : # ctype : 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
                data.append(float(sheet1.cell(i,j).value) * sf)
            else:
                data.append(sheet1.cell(i,j).value)
        list_data.append(data)
    datamatrix = np.array(list_data)    
    return datamatrix 

# # Test    
# import os
# path = 'C:\\Users\\68075\\OneDrive - The University of Texas at Austin\\Desktop\\Grid Resilience\\Base_pipeline'
# file_name = 'base.xlsx'
# file_path=os.path.join(path,file_name)
# matrix = read_excel(file_path,'Line',0,1,0,1,1.0)
# print(matrix)

# Loads = read_excel('Base_Parameters.xlsx', 'Load Parameters', 4, 5, 1, 3, 1.0)
# target_load = Loads[1,:]
# print(Loads)
# print(target_load)
