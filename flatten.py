"""
This script is intended to flatten an excel file that had a crosstab approach
with each option from a survey response a separate column. This led to about 34 columns for the 17 options.

Tableau visualization wants row-based data table with each question response a different column.

It would probably take 10 minutes to manually recreate, but instead I wanted to learn and practice python, so this script.

Output is a output.csv file with each respondant repeated with 5 columns:
Person	C/I/O	Role-Level	FunctionLevel	FunctionResp	FunctionSupport
Person=name of responder
C/I/O=center of responder
Role-Level=what kind of position-center, office, advisor
FunctionLevel=if the center has this center, values are "B" if both responsible and supporting, "R" if responsible, "S" if supporting, "0" for none.
"""

import sys
import openpyxl
import csv

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils import column_index_from_string

wb = load_workbook('analysis.xlsx')

ws = wb.get_sheet_by_name('Data-Responses-FlatFunctions')



numRows = 17 #16 responses plus 1 header
print 'numRows=' + str(numRows)

numCols = 70 #only the function responses and a few pre columns
print 'numCols=' + str(numCols)

numFunctions = 32 #32 different functions that could be selected for Q4 and Q5

output = open('output.csv',"wb")
writer = csv.writer(output,delimiter=',',quotechar='"',quoting=csv.QUOTE_ALL)
writer.writerow(['Person','C/I/O','Role-Level','Function','Auth','Support'])

for row in ws['A2':'BR17']:
	for cell in row:
		headerColName = ws[cell.column + '1'].value	
		cellValue = '0'
		if headerColName.startswith('Q4'):
			responsibleFunction = '0'
			supportsFunction = '0'
			if cell.value is not None:#if it has any value, then person is resposible
				responsibleFunction ='R'
				cellValue = 'R'

			#now check the equivalent column for supporting
			q4ColNumber = column_index_from_string(cell.column)
			q5ColLetter = get_column_letter(q4ColNumber + numFunctions)
			functionName = ws[cell.column + '1'].value[4:]
			print 'q4ColNumber=' + str(q4ColNumber)
			print 'q5ColLetter=' + q5ColLetter
			if ws[q5ColLetter + str(cell.row)].value is not None: #if it has any value, then person is responsible
				supportsFunction = 'S'
				print 'supporting question=' + ws[q5ColLetter + str(cell.row)].value
			
			print 'responsibleFunction= ' + responsibleFunction
			print 'supportsFunction= ' + supportsFunction
			print 'functionName= ' + functionName


			writer.writerow([ws['A' + str(cell.row)].value,ws['B' + str(cell.row)].value,ws['C' + str(cell.row)].value,ws[cell.column + '1'].value[3:],cellValue])
		elif headerColName.startswith('Q5'):
			if cell.value is not None:
                                cellValue = 'S'
                        print cellValue + ws[cell.column + '1'].value[3:]
			#for supporting, need an extra space
			writer.writerow([ws['A' + str(cell.row)].value,ws['B' + str(cell.row)].value,ws['C' + str(cell.row)].value,ws[cell.column + '1'].value[3:],'',cellValue])

output.close()
