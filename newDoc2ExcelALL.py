############### Data from Docx
import docx2txt, re
import sys
import openpyxl

wb = openpyxl.Workbook()
sheet = wb['Sheet']
sheet['A1'] = 'FIRST'
sheet['B1'] = 'LAST'
sheet['C1'] = 'ADDRESS'
sheet['D1'] = 'CITY'
sheet['E1'] = 'STATEZIP'
whatDocument = int(input("\n1) Montco MF \n2) Bucks MF \n3) Montco Civil\n4) Bucks Civil\n\nInput corresponding document number: "))
if (whatDocument == 1):
    document = 'MontcoMF.docx'
elif (whatDocument == 2):
    document = 'BucksMF.docx'
elif (whatDocument == 3):
    document = 'MontcoCivil.docx'
elif (whatDocument == 4):
    document = 'BucksCivil.docx'
bigData = docx2txt.process('/Users/garcgabe/Downloads/NewAutomation/' + str(document)) #whole doc


exceptions1Reg = re.compile(r'Select\s(.*)\sAKA')
exceptions1 = exceptions1Reg.findall(str(bigData))

exceptions2Reg = re.compile(r'Select\s(.*)\sA/K/A')
exceptions2 = exceptions2Reg.findall(str(bigData))

aptexceptionsReg = re.compile(r'Defendants\s\nSelect\t(.*)\n(.*)\n(.*)\sUNITED STATES')
apts = aptexceptionsReg.findall(str(bigData))


### Filter out Internal Use Only messages --> so they're things u may actually have to change
reals = []
for bit in exceptions1:
    chunk = str(bit)
    if 'Internal Use Only' in chunk:
        x = 1
    elif '2021' in chunk:
        x = 2
    else:
        reals.append(bit)
for bit in exceptions2:
    chunk = str(bit)
    if 'Internal Use Only' in chunk:
        x = 1
    elif '2021' in chunk:
        x = 2
    else:
        reals.append(bit)
for bit in apts:
    chunk = str(bit)
    if 'Internal Use Only' in chunk:
        x = 1
    elif '2021' in chunk:
        x = 2
    else:
        reals.append(bit[0])

print('Some exceptions regarding APT and SUITE numbers or names in LAST, FIRST order may be: \n\n' + str(reals))

wholesectionReg = re.compile(r'Defendants\s\nSelect\t(.*)\n(.*)\sUNITED STATES')
sections = wholesectionReg.findall(str(bigData))

for i in range(1,len(sections)+1):
    A = sections[i-1]
    NameAddy = str(A).split(r'\t')
    fullname = NameAddy[0]
    if 'AKA' in fullname:
        print('Problem with ' + str(fullname) + '. Edit in Excel.')
    if 'A/K/A' in fullname:
        print('Problem with ' + str(fullname) + '. Edit in Excel.')
    if ',' in fullname:
        print('')
    else:
        print('No comma in ' + fullname + ' may be a business. Find name in DOCX and replace name with USE, DO NOT. or format as LAST, FIRST name.')
    SplitName = str(fullname).split(', ')
    lastname = SplitName[0]
    firstname = SplitName[1]
    BigAddy = NameAddy[1]
    SplitAddy = str(BigAddy).split(', ')
    addy = SplitAddy[0]
    city = SplitAddy[1]
    if 'UNKNOWN' in SplitAddy:
        print('Incomplete info for ' + fullname)
    statezip = SplitAddy[2]
    if '(' in firstname:
        print('Trim ' + str(lastname))
    if 'AKA' in str(addy):
        print('Check Address for duplicate.')
    if 'A/K/A' in str(addy):
        print('Check Address for duplicate.')
    statezip = statezip[0:8]

    sheet['A'+str(i+1)] = "=Proper(\"" + str(firstname) + "\")"
    sheet['B'+str(i+1)] = "=Proper(\"" + str(lastname[2:]) + "\")"
    sheet['C'+str(i+1)] = "=Proper(\"" + str(addy[0:len(addy)-1]) + "\")"
    sheet['D'+str(i+1)] = "=Proper(\"" + str(city[1:]) + "\")"
    sheet['E'+str(i+1)] = str(statezip)


    wb.save('scraped_sheet.xlsx')
# person 2


sys.exit("All good dog")
