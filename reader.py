# from collections import OrderedDict
import openpyxl
# import pickle

print'Opening armor book...'
wb = openpyxl.load_workbook('/tmp/mozilla_busby0/sr_armor .xlsx')
sheet = wb.get_sheet_by_name('Sheet1')
armors = {'name': [], 'rating': [], 'capacity': [], 'avail': [], 'cost': []}
print 'Reading rows'
for row in range(2, sheet.max_row + 1):
    armors['name'] = str(sheet['A' + str(row)].value.strip('*'))
    armors['rating'] = str(sheet['B' + str(row)].value).strip('.0')
    armors['capacity'] = str(sheet['C' + str(row)].value).strip('.0')
    armors['avail'] = str(sheet['D' + str(row)].value).strip('.0')
    armors['cost'] = sheet['E' + str(row)].value[:-1].replace(',', '')
    print "{0:s}".format(armors['name']), armors["rating"], armors['capacity'], armors['avail'], armors['cost']

# print "Armors(%s)" % armors{'name'}
