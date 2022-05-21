import openpyxl



workbook = openpyxl.load_workbook("tester.xlsx")
sheet = workbook.active
dictionary = {}



val=sheet.cell(row=1, column=1)
obj = val.value
stripper = obj.replace('"', '')
individuals = stripper.split(",")



for temp in individuals:
    finalSplit = temp.split(':')
    key = finalSplit[0]
    val = finalSplit[1]
    dictionary[key]=val


print(dictionary["appname"])
