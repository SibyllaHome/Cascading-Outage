
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


def MatchingElement_IEEE39(name = '0', bus = '0', cub = '0'):
    import Excel
    row = '0'
    value  = '0'    
    type = 'Notfound'
    data = Excel.read_excel('IEEE39_Parameters.xlsx', 'Bus Parameters', 4, 78, 1, 7, 1.0)
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

# print(MatchingElement_IEEE39(bus = '16', cub = '7'))

def MatchingLine_IEEE39(num1, num2):
    import Excel
    value  = '0'    
    data = Excel.read_excel('IEEE39_Parameters.xlsx', 'Line Parameters', 4, 49, 1, 5, 1.0)
    if num1 == '01' and num2 == '01' :
        value = '1'
    if num1 == '02' and num2 == '39' :
        value = '2'
    if num1 == '39' and num2 == '02' :
        value = '2'
    if num1 == '17' and num2 == '17' :
        value = '30'
    if num1 == '18' and num2 == '27' :
        value = '31'
    if num1 == '27' and num2 == '18' :
        value = '31'
    else :
        for i in range(data.shape[1]):
            if num1 == data[2, i] :
                if num2 == data[3, i] : 
                    value = data[1, i]
                    break
            if num1 == data[3, i] :
                if num2 == data[2, i] :
                    value = data[1, i]
                    break
    return value

def MatchingGen_IEEE39(num):
    import Excel
    data = Excel.read_excel('IEEE39_Parameters.xlsx', 'Generator Parameters', 3, 12, 1, 2, 1.0)
    for i in range(data.shape[1]):
        if num == data[0, i] :
            value = data[1, i]
            break
        if num == data[1, i] :
            value = data[0, i]
            break
    return value  