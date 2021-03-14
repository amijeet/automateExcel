import xlrd
import xlsxwriter

location = "./bomFile.xls"
wb = xlrd.open_workbook(location)
sheet = wb.sheet_by_index(0)

data = []
i = 0

element = sheet.cell_value(1, 0)

# CONSTRUCT THE DATA ARRAY
while( i < sheet.nrows-1 ):
    data.append([])
    data[i].append(sheet.cell_value(i+1, 0))
    data[i].append(sheet.cell_value(i+1, 1))
    data[i].append(sheet.cell_value(i+1, 2))
    data[i].append(sheet.cell_value(i+1, 3))
    data[i].append(sheet.cell_value(i+1, 4))
    i += 1
    element = sheet.cell_value(i, 0)

# CONVERT LEVEL TO A NUMERIC STRING
for arr in data:
    temp = arr[1]
    num = ""
    for x in temp:
        if( x.isnumeric() ):
            num += x
    arr[1] = num

# PRINTING THE DATA ARRAY
for x in data:
    print(x)

# LIST OF LEVEL 1 GOODS
level1 = []
compare = ""
for arr in data:
    if( compare != arr[0] ):
        level1.append(arr[0])
        compare = arr[0]

# MAX LEVEL OF LEVEL1 GOODS
maxLevel = []
count = 0
max = -1
for arr in data:
    if( level1[count] == arr[0] ):
        if( int(arr[1]) > max ):
            max = int(arr[1])
    else:
        maxLevel.append(max)
        max = -1
        count += 1
maxLevel.append(max)

# PRINT LEVEL1 ANS LEVEL1 MAX
for x in level1:
    print( x )
for x in maxLevel:
    print(x)

count = len(data)
current = data[0][0]
# MAIN ALGO for LEVEL1
for name in level1:
    count = 0
    workbook = xlsxwriter.Workbook(name + '.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, "Finished Good List")
    worksheet.write(1, 0, "#")
    worksheet.write(1, 1, "Item Description")
    worksheet.write(1, 2, "Quantity")
    worksheet.write(1, 3, "Unit")
    worksheet.write(2, 0, "1")
    worksheet.write(2, 1, name)
    worksheet.write(2, 2, "1")
    worksheet.write(2, 3, "Pc")
    worksheet.write(3, 0, "End of FG")
    worksheet.write(4, 0, "Raw Material List")
    worksheet.write(5, 0, "#")
    worksheet.write(5, 1, "Item Description")
    worksheet.write(5, 2, "Quantity")
    worksheet.write(5, 3, "Unit")
    for arr in data:
        if( arr[0] == name and arr[1] == "1" ):
            worksheet.write(6+count, 0, str(count+1))
            worksheet.write(6+count, 1, arr[2])
            worksheet.write(6+count, 2, arr[3])
            worksheet.write(6+count, 3, arr[4])
            count += 1
    workbook.close()

# ALGO FOR LOWER LEVEL ITEMS
for i in range(len(data)):
    if( int(data[i][1]) != 1 ):
        if( int(data[i-1][1]) == int(data[i][1])-1 ):
            num = int(data[i][1])
            count = 0
            workbook = xlsxwriter.Workbook(data[i-1][2] + '.xlsx')
            worksheet = workbook.add_worksheet()
            worksheet.write(0, 0, "Finished Good List")
            worksheet.write(1, 0, "#")
            worksheet.write(1, 1, "Item Description")
            worksheet.write(1, 2, "Quantity")
            worksheet.write(1, 3, "Unit")
            worksheet.write(2, 0, "1")
            worksheet.write(2, 1, data[i-1][2])
            worksheet.write(2, 2, "1")
            worksheet.write(2, 3, "Pc")
            worksheet.write(3, 0, "End of FG")
            worksheet.write(4, 0, "Raw Material List")
            worksheet.write(5, 0, "#")
            worksheet.write(5, 1, "Item Description")
            worksheet.write(5, 2, "Quantity")
            worksheet.write(5, 3, "Unit")

            while( int(data[i][1]) == num ):
                worksheet.write(6+count, 0, str(count+1))
                worksheet.write(6+count, 1, data[i][2])
                worksheet.write(6+count, 2, data[i][3])
                worksheet.write(6+count, 3, data[i][4])
                count += 1
                i += 1
            worksheet.write(6+count, 0, "End of RM")
            workbook.close()
