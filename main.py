import xlrd
from collections import OrderedDict
import simplejson as json
 
# Open the workbook and select the first worksheet
workBook = xlrd.open_workbook('quiz.xls')
workSheet = workBook.sheet_by_index(0)
 
# List to hold dictionaries
quiz_list = []
 
# Iterate through each row in worksheet and fetch values into dict
for rownum in range(1, workSheet.nrows):
    quiz = OrderedDict()
    row_values = workSheet.row_values(rownum)
    quiz['Question'] = row_values[0]
    quiz['Answer'] = row_values[1]
    quiz['type'] = row_values[2]
    
 
    quiz_list.append(quiz)
 
# Serialize the list of dicts to JSON
j = json.dumps(quiz_list)
 
# Write to file
with open('data.json', 'w') as f:
    f.write(j)