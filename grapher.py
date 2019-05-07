import xlwt
from xlrd import open_workbook
from xlutils.copy import copy
import PyPDF2, re
import pandas

#open pdf file and get data
whichPdf = input('what file?')
pdfFileObj = open(whichPdf+'.pdf', 'rb')
pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
pageObj = pdfReader.getPage(0)
text = pageObj.extractText()
text = text.replace('\n','')

with open("Output.txt", "w") as text_file:
    text_file.write(text)

#extracts data into lists
regex = (r'\d+ (...) (\d{3})')
classes = re.findall(regex, text)
if classes == []:
    regex = (r'\d+(...) (\d{3})')
    classes = re.findall(regex, text)
regex2 = r'(..)(\d:\d\d ..) - (\d:\d\d ..)'
classes2 = re.findall(regex2, text)
regex3 = r'((M|W|F|T|Th| M| T| W| Th| F|M |T |W |Th |F )+)(\d+:\d\d ..) - (\d+:\d\d ..)'
Times = re.findall(regex3, text)
regex4 = r'Tempe (.+ \d\d\d)'
locations = re.findall(regex4,text )

classesSummed = []
for each in classes:
    classesSummed.append(each[0]+each[1])
finalList=[]
print(classesSummed)
for each in Times:
    className = classesSummed[Times.index(each)]
    each+=(className,)
    finalList.append(each)
    print(each)

locationList = []
for each in locations:
    locationList.append(each)
    print(each)



days = ['M', 'T', 'W', 'Th', 'F']




#creates graph outline
def createGraph(name):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('test sheet')
    extension = ' AM'
    minute = 0
    hour = 7
    time=''
    timeList = []
    rowCounter =1;
    while time!='8:00 PM':
        if (hour==12 and extension==' AM'):
            extension = ' PM'
        if (hour==12 and minute==60):
            hour=1
            minute = minute-60
        if minute==60:
            minute = minute-60
            hour=hour+1
        time = str(hour)+':'+str(minute)
        time = time+extension
        if len(str(minute).replace(' ', '')) ==1:
            if '5 ' in time:
                time = time.replace('5 ', '05 ')
            if '0 ' in time:
                time = time.replace('0 ', '00 ')
        if time == '12:00 AM':
            time = '12:00 PM'
        timeList.append(time)
        ws.write(rowCounter,0,time)
        rowCounter=rowCounter+1
        minute=minute+5
    for eachDay in days:
        ws.write(0, days.index(eachDay)+1, eachDay)

    wb.save(name)

filename = whichPdf+'Schedule.xls'
createGraph(filename)



#reads data and sets up regex to find row
df = pandas.read_excel(filename)
regexRow = re.compile(r'(\d+:\d\d ..)')

def findTimeNum(num):
    for each in range(158):
        mo = regexRow.search(str(df.iloc[each]))
        if mo is not None:
            if mo.group(0) == num:
                return each


def findDayNum(num):
    for each in days:
        if num == each:
            return days.index(each)


#reopens file to edit
rb = open_workbook(filename)
wb = copy(rb)
s = wb.get_sheet(0)


def plot(time):
    row = findTimeNum(time[2])+1
    time = list(time)
    time[1] = time[1].replace(' ','')
    dayList = time[0].split(' ')
    column = findDayNum(time[1])+1
    finishRow = findTimeNum(time[3])+1
    for each in range(row, finishRow+1):
        for eachday in dayList:
            currentday = findDayNum(eachday)+1
            s.write(each, currentday, time[4])

for each in finalList:
    try:
        plot(each)
    except Exception as e:
        print(e, each)

wb.save(filename)
