'''
Project Name:Excel Divider
Author: Yuzhao Liu
Created Date: 1/15/2019
'''


'''
Note: For now, all the required parameter for the function "read_original_files" are hard coded as file name, total lines, and liens per file
we can change this later if needed

Note: Please delete all other redundent files before running!!!!!!!!!!!
'''



import xlrd
import xlwt



def set_header(ws):
    '''
    set up the header for each col
    '''
    ws.write(0,0,"D-U-N-S Number")
    ws.write(0,1,"Parent Company Name")
    


def read_original_files(old_file, old_lines, lines_limit):
    '''
    old_file is the dir of the original file
    old_lines is how many lines need to be processed
    lines_limit is how many line per new file
    
    '''
    original_file_location = old_file
    original_file = xlrd.open_workbook(original_file_location)
    sheet = original_file.sheet_by_index(0) ##hard code as the first sheet
    num_of_new_files = old_lines//lines_limit
    lines_in_last_file = old_lines%lines_limit

    row_count = 1
    
    for file in range(num_of_new_files):
        wb = xlwt.Workbook()
        ws = wb.add_sheet("sheet1")
        
        set_header(ws)
        
        for line in range(lines_limit):
            col0 = sheet.cell_value(row_count,0)
            col1 = sheet.cell_value(row_count,1)
            
            ws.write(line+1,0,col0)
            ws.write(line+1,1,col1)
            
            row_count+=1
            
        new_file_name = str(file+1)
        wb.save(f"{new_file_name}.xls")

    
    if lines_in_last_file > 0:
        wb = xlwt.Workbook()
        ws = wb.add_sheet("sheet1")

        set_header(ws)
        
        for line in range(lines_in_last_file):
            col0 = sheet.cell_value(row_count,0)
            col1 = sheet.cell_value(row_count,1)
            
            ws.write(line+1,0,col0)
            ws.write(line+1,1,col1)
            
            row_count+=1
            
        new_file_name = str(num_of_new_files+1)
        wb.save(f"{new_file_name}.xls")



if __name__ == "__main__":
##    valid = 1
##    total_lines = input("Please Enter the Total Lines Num: ")
##    try:
##        int(total_lines)
##    except:
##        print("Please Enter A Valid Num. Run the Whole Program Again Please")
##        valid = 0
##    lines_each_file = input("Please Enter the Lines for Each File: ")
##    
##    try:
##        int(lines_each_file)
##    except:
##        print("Please Enter A Valid Num. Run the Whole Program Again Please")
##        valid = 0
##    if valid == 1:



    total_lines = 2373
    lines_each_file = 150
    
    file = r"C:\Users\lyz95\Desktop\excel_divider\Parent_companies.xls"
    read_original_files(file, total_lines, lines_each_file)




















