# def MatchingLine_Base(num):
#     import Excel
#     value  = '0'    
#     data = Excel.read_excel('Base_Parameters.xlsx', 'Line Parameters', 4, 6, 1, 6, 1.0)
#     for i in range(data.shape[1]):
#         if num == data[1, i] :
#             value = data[0, i]
#             break
#         if num == data[0, i] :
#             value = data[1, i]
#             break
#     return value

'''
MatchingLine_Base: match lines in Base.pfd
Input: 
bus - the bus of element
cub - the cub of line
Output:
type - the type of element
value - the element name of element 
row - the index of element in corresponding excel file
'''
from xlrd.book import Name


def MatchingElement_Base(name = '0', bus = '0', cub = '0'):
    import Excel
    row = '0'
    value  = '0'    
    type = 'Notfound'
    data = Excel.read_excel('Base_Parameters.xlsx', 'Bus Parameters', 4, 37, 1, 7, 1.0)
    for i in range(data.shape[1]):       
        if ((bus == data[2, i]) & (cub == data[4, i])) | ((bus == data[3, i]) & (cub == data[5, i])):
            type = data[6,i]
            value = data[1, i]
            row = data[0, i]
            break
        if name == data[1, i]:
            type = data[6,i]
            value = name
            row = data[0, i]
            break
    return type, value, row

# # test
# print(MatchingLine_Base(name = '1'))

# def MatchingGen_Base(num):
#     import Excel
#     data = Excel.read_excel('Base_Parameters.xlsx', 'Generator Parameters', 4, 6, 1, 3, 1.0)
#     for i in range(data.shape[1]):
#         if num == data[0, i] :
#             value = data[1, i]
#             break
#         if num == data[1, i] :
#             value = data[0, i]
#             break
#     return value    