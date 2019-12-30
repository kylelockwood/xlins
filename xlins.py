#! python3
# Inserts rows in the excel active sheet
# Usage: xlins <file name.type> <sheet> <row> <many>

import sys, openpyxl, os, glob

# Find open excel file in current working directory        

sheetName = ""
xlfiles = []
fileName = str(sys.argv[1])

# Insure proper input
if fileName == "list":
    for files in glob.glob('*.xls*'):
        xlfiles.append(files)
    if len(xlfiles) == 0:
        sys.exit(f'No excel files found in {str(os.getcwd())}')
    sys.exit(f'Excel files in {str(os.getcwd())} :\n{xlfiles}')

if not sys.argv[2] == "list":
    try:
        sheetName = str(sys.argv[2])   
        rowNum = int(sys.argv[3])
        insNum = int(sys.argv[4])
    except:
        sys.exit('Error: Incorrect usage : xlins <file name.type> <sheet name> <row> <many>')
    if rowNum < 1:
        sys.exit('Error: <row> must be an integer greater than 0')
    if insNum < 1:
        sys.exit('Error: <many> must be an integer greater than 0')
try:
    fullFileName = str(os.getcwd() + '\\' + fileName)
    wb = openpyxl.load_workbook(fullFileName)
except:
    sys.exit(f'Error: could not find {fileName} in current directory.')

# Open workbook and sheet
try:
    sheet = wb[sheetName]
except:
    if not sys.argv[2] == "list":
        print(f'Sheet "{sheetName}" not found in {fileName}')
    sys.exit(f'Sheet names in {fileName} : {wb.sheetnames}')

# Insert rows
for i in range(0, insNum):
    sheet.insert_rows(rowNum)

# Copy Data
copyRow = []
copyData = []
for i in range(rowNum,sheet.max_row + 1):
    for j in range(1, sheet.max_column + 1):
        copyData.append(sheet.cell(row = i, column = j).value)
    copyRow.append(copyData)
    copyData = []

print(copyRow)

# Paste data
pasteRow = rowNum + insNum - 1
k = 0
for i in range(pasteRow, pasteRow + insNum - 1):
    for j in range(0, sheet.max_column):
        print(f'row : {i}, col : {j +1}, data : {copyRow[k][j]}')
        sheet.cell(row = i, column = j + 1).value = copyRow[k][j]
    k += 1

print(f'Inserted {insNum} rows at row {rowNum} in {fileName} {sheetName}')
wb.save(fullFileName)