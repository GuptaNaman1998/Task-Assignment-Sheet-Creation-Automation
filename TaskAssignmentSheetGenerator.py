from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border,Font, Alignment,Side
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter
import warnings
import datetime
import time

def AddSheet(ws,week,POC):
    # Filling the Header with the values
    ws.append(['Shift','Engineer Name','Case','Type','Case','Type','Case','Type','Case','Type'])
    x=week.keys()

    for i in x:
        # Filling the cells with the values from the Dictionary and checking for POCs
        if i in POC:
            ws.append(["POC : "+week[i],i])
        else:
            ws.append([week[i],i])

    # Format the cell colour to be as desired
    OrangeFill = PatternFill(start_color='FFC000',end_color='FFC000',fill_type='solid')
    BlueFill = PatternFill(start_color='00B0F0',end_color='00B0F0',fill_type='solid')
    BlackFill = PatternFill(start_color='000000',end_color='000000',fill_type='solid')
    GreenFill = PatternFill(start_color='33CC33',end_color='33CC33',fill_type='solid')

    # Format the cell border
    ThinBorder = Border(left=Side(style='thick'),right=Side(style='thick'),top=Side(style='thick'),bottom=Side(style='thick'))

    # Format the Font size type and appearance
    BlackFont= Font(name='Cambria',size=11,bold=True,italic=True)
    WhiteFont = Font(name='Cambria',size=11,bold=True,italic=True,color='FFFFFF')

    # Format the text alignment in the cells
    Al= Alignment(horizontal='center',vertical='center')

    # Applying the Styles & formats to the cells
    for col in range(1, ws.max_column + 1):
        cell_header = ws.cell(1, col)
        # print(cell_header.value)
        cell_header.fill = GreenFill
        cell_header.border = ThinBorder
        cell_header.font = BlackFont
        cell_header.alignment = Al
    for cell in ws['A2:A{}'.format(ws.max_row)]:
        cell[0].fill = OrangeFill
        cell[0].border = ThinBorder
        cell[0].font = BlackFont
        cell[0].alignment = Al
    for cell in ws['B2:B{}'.format(ws.max_row)]:
        cell[0].fill = OrangeFill
        cell[0].border = ThinBorder
        cell[0].font = BlackFont
        cell[0].alignment = Al
    for cell in ws['D2:D{}'.format(ws.max_row)]: 
        cell[0].fill = BlueFill
        cell[0].border = ThinBorder
        cell[0].font = BlackFont
        cell[0].alignment = Al
    for cell in ws['F2:F{}'.format(ws.max_row)]: 
        cell[0].fill = BlueFill
        cell[0].border = ThinBorder
        cell[0].font = BlackFont
        cell[0].alignment = Al
    for cell in ws['H2:H{}'.format(ws.max_row)]: 
        cell[0].fill = BlueFill
        cell[0].border = ThinBorder
        cell[0].font = BlackFont
        cell[0].alignment = Al
    for cell in ws['J2:J{}'.format(ws.max_row)]:
        cell[0].fill = BlueFill
        cell[0].border = ThinBorder
        cell[0].font = BlackFont
        cell[0].alignment = Al
    for cell in ws['I2:I{}'.format(ws.max_row)]:
        cell[0].border = ThinBorder
        cell[0].font = BlackFont
        cell[0].alignment = Al
    for cell in ws['C2:C{}'.format(ws.max_row)]:
        cell[0].border = ThinBorder
        cell[0].font = BlackFont
        cell[0].alignment = Al
    for cell in ws['E2:E{}'.format(ws.max_row)]:
        cell[0].border = ThinBorder
        cell[0].font = BlackFont
        cell[0].alignment = Al
    for cell in ws['G2:G{}'.format(ws.max_row)]:
        cell[0].border = ThinBorder
        cell[0].font = BlackFont
        cell[0].alignment = Al
        
    # Marking the POC with different cell colour and font colour
    for row in range(1, ws.max_row + 1):
        cell_header = ws.cell(row, 1)
        if "POC" in cell_header.value:
            cell_header.fill = BlackFill
            cell_header.font = WhiteFont
            cell_header = ws.cell(row, 2)
            cell_header.fill = BlackFill
            cell_header.font = WhiteFont 

    # Setting the Column width
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 26           
    ws.column_dimensions['C'].width = 12            
    ws.column_dimensions['D'].width = 12            
    ws.column_dimensions['E'].width = 12            
    ws.column_dimensions['F'].width = 12            
    ws.column_dimensions['G'].width = 12            
    ws.column_dimensions['H'].width = 12            
    ws.column_dimensions['I'].width = 12            
    ws.column_dimensions['J'].width = 12

def ExtractData(file):
    '''
        This code reads the Excel sheet and cleans or sorts out the data we require.
        After the sorting and filtering of the data it returns a list of Required Data data.
        
        Data We Have:
            Name of the Engineers
            Days of the month
            Days of the week
            Shift Letter specifying the shift of that engineer
            POC Details
            Days in each shift per engineer
            PH & PL
            OnCall Details
            
        Data We Require To be returned:
            Name of the Engineers
            Days of the month
            Days of the week
            Shift Letter specifying the shift of that engineer
            PH & PL
    '''
    # Asking for the file name to be read
    
    wb = load_workbook(file+'.xlsx', data_only = True)
    ws = wb.active
    sh = wb[ws.title]
    temp=[]
    for row in sh.rows:
        tempr=[]
        for cell in row:
            if cell.value=='A - (6:00 AM - 3:00 PM)':
                return temp[1:-1]
            tempr.append(cell.value)
        temp.append(tempr)
        
def roster(flag,temp):
    '''
        This code takes a list of Required Data as input and returns a dictionary.
        
        Arguments/Parameters:
            Flag - It hols the value of the day of the month on the First monday
            temp - It is the list of the required filtered data
            
        Returned Variable:
            A dictionary sorted based on the values of the dictionary, with Key as the engineer name & Value as their shift letter
    '''
    week=[]
    if flag <= 27:
        x=5
    else:
        x=32-flag
    for ele in range(x):
        week1={}
        for i in temp[2:]:
            shifts = i[flag+ele]
            if shifts == 'L' :
                shifts='Leave'
            if shifts == None:
                shifts='NA'
            week1[i[0]] = shifts 
        week.append({k: v for k, v in sorted(week1.items(), key=lambda item: item[1])})
    return week

def createExcel(week,flag,fname):
    '''
        This code creates the Excel sheet as an output.
        
        Arguments/Parameters:
            Flag - It hols the value of the day of the month on the First monday
            Week - It is the Dictionary of the required filtered data
            
        Returned Variable:
            A dictionary sorted based on the values of the dictionary, with Key as the engineer name & Value as their shift letter
    '''
    wb = Workbook()
    
    # Asking for POC details from the user
    Val=input("Please enter the names of POCs followed by a ',' :")
    POC=Val.split(",")
    
    # Logic for days of the last week 
    if flag <= 27:
        x=5
    else:
        x=32-flag
        
    # Looping the function call over the number of days
    for i in range(x):
        ws = wb.create_sheet(str(flag+i), i)
        AddSheet(ws,week[i],POC)
        
    d = time.strptime(fname, "%d %b %Y")  
    weekNumber = datetime.date(d.tm_year,d.tm_mon,d.tm_mday).strftime("%V")
    # Saving the output in the predefined folder
    wb.save('Output/Week-'+weekNumber+'.xlsx')
    
if __name__ == "__main__":
    warnings.simplefilter("ignore")
    print("*Note: entered file name should be in the month-year format eg. August-2021")
    file = input("Enter the file name to read (without extension):")
    val = ExtractData(file)
    c = file.split("-")
    temp=[]
    for i in val:
        temp.append(i[1:33])
        
    for i in range(len(temp[0])):
        if temp[0][i]=='Mon':
            flag=i
            break
    
    while flag<32:
        fname = str(flag)+" "+c[0][:3]+" "+c[1]
        week1=roster(flag,temp)
        createExcel(week1,flag,fname)
        flag+=7
        
"""
Exceptions/Lowlights:
        
        1. February month the flag threshold will need to be changed
        2. For all the months only whole weeks are being considered, Basically the code finds the first cell where monday appears and considers it as day 1.
        Which will cause a week of missing roster.
        3. Have to add the leaves, Public Holidays and training part manually
        4. For every month's roster the code will ignore the days before the first monday but will consider the last week even if it is just a monday.
        This will get an errored roster as if a person was on leave on monday the code will add leave for the whole week.
        
Next Implementation Steps:

        1. Convert the file name preset value to user input.
        2. Convert this script to either a GUI application or an executable ".exe" file.
        3. Need to add naming convention as we need the file to be named week whatever the count is.
        
Subjeactive Addition:

        1. Currently the naming convention of case assignment file is Week+ " Count of the week " which leads to sorting issues.
        What if we create a folder heirarcy wherein you don't have to find which file has the highest week written on it, Just have to scroll to the end of the folder.
        2. POC Addition/Incusion in the case assignment sheet generated as an output.
"""
