from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border,Font, Alignment,Side
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter

def ExtractData():
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
    wb = load_workbook('July-21.xlsx', data_only = True)
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
    week1={}
    for i in temp[2:]:
        shifts=list(set(i[flag:flag+5]))
        if 'L' in shifts:
            shifts.remove('L')
        if len(shifts)==0 :
            shifts=['Leave']
        if shifts[0]==None:
            shifts=['na']
        week1[i[0]]=shifts[0]
    return {k: v for k, v in sorted(week1.items(), key=lambda item: item[1])}

def createExcel(week,flag):
    '''
        This code creates the Excel sheet as an output.
        
        Arguments/Parameters:
            Flag - It hols the value of the day of the month on the First monday
            Week - It is the Dictionary of the required filtered data
            
        Returned Variable:
            A dictionary sorted based on the values of the dictionary, with Key as the engineer name & Value as their shift letter
    '''
    wb = Workbook()

    ws = wb.active
    ws.title=str(flag)

    # Saving the Keys from the dictionary as a list to later use it to iterate through the dictionary
    x=week.keys()
    
    # Heading or the column title for the Sheet
    ws.append(['Shift','Engineer Name','Case','Type','Case','Type','Case','Type','Case','Type'])
    
    for i in x:
        # Filling the cells with the values from the Dictionary
        ws.append([week[i],i])

    # Format the cell colour to be orange
    redFill = PatternFill(start_color='FFFFC000',end_color='FFFFC000',fill_type='solid')
    
    # Format the cell colour to be blue 
    blueFill = PatternFill(start_color='00B0F0',end_color='00B0F0',fill_type='solid')
    
    # Format the cell border
    thin_border = Border(left=Side(style='thick'),right=Side(style='thick'),top=Side(style='thick'),bottom=Side(style='thick'))
    
    # Format the Font size type and appearance
    font= Font(name='Cambria',size=11,bold=True,italic=True)
    
    # Format the text alignment in the cells
    Al= Alignment(horizontal='center',vertical='center')
    
    # Apply the color formats to the cells
    for col in range(1, ws.max_column + 1):
        cell_header = ws.cell(1, col)
        cell_header.fill = PatternFill(start_color='33CC33', end_color='33CC33', fill_type="solid")
    for cell in ws['A2:A{}'.format(ws.max_row)]:
        cell[0].fill = redFill
    for cell in ws['B2:B{}'.format(ws.max_row)]:
        cell[0].fill = redFill
    for cell in ws['D2:D{}'.format(ws.max_row)]: 
        cell[0].fill = blueFill
    for cell in ws['F2:F{}'.format(ws.max_row)]: 
        cell[0].fill = blueFill
    for cell in ws['H2:H{}'.format(ws.max_row)]: 
        cell[0].fill = blueFill
    for cell in ws['J2:J{}'.format(ws.max_row)]:
        cell[0].fill = blueFill
        
    # Apply the font , border and alignment
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter # Get the column name
        for cell in col:
            cell.border = thin_border
            cell.font = font
            cell.alignment = Al
            if cell.coordinate in ws.merged_cells: # not check merge_cells
                continue
            try: # Necessary to avoid error on empty cells
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column].width = 10  
        
    # Save the file #FFC000  #33CC33
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 25  
    
    wb.copy_worksheet(ws).title=str(flag+1)
    wb.copy_worksheet(ws).title=str(flag+2)
    wb.copy_worksheet(ws).title=str(flag+3)
    wb.copy_worksheet(ws).title=str(flag+4)
    
    wb.save('Output/Monday-'+str(flag)+'th'+'.xlsx')
    
if __name__ == "__main__":
    val=ExtractData()
    temp=[]
    for i in val:
        temp.append(i[1:33])
    
    for i in range(len(temp[0])):
        if temp[0][i]=='Mon':
            flag=i
            break
    # print(flag)
    
    # Loop the call to create the excel sheet for all the mondays
    while flag<32:            
        week1=roster(flag,temp)
        createExcel(week1,flag)
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
