#-*- coding: utf8 -*-
import os
import xlrd
import xlwt

# input directory path and return a handle list contains all opened excel file  
# the path should use '\\' rather than '\'
def open_all_files(path):
    fname_list = []
    fhandle_list = []
    for fname in os.listdir(path):
        fname_list.append( fname)
        fhandle_list.append( xlrd.open_workbook(path + fname) )
    if ( len(fname_list) != len(fhandle_list)):
        print ("Error: Could not open all files in specified directory")
    return (fname_list, fhandle_list)

def open_merge_sheet( fhandle_list) :
    sheet_list = []
    for fhandle in fhandle_list:
        try:
            sheet_list.append( fhandle.sheet_by_name('Sheet2'))
        except :
            print("Sheet1 was open instead, because the miss of Sheet2")
            sheet_list.append( fhandle.sheet_by_name('Sheet1'))
    return sheet_list
def merge_sheet( sheet_list):
    title =['list type	List', 'snoop table', 'List Index',	'Total', 'Assign to', 'status',  'Description', 'Waveform & log dir:', 'pending on']
    wbk = xlwt.Workbook(encoding='utf-8')  
    sheet_w = wbk.add_sheet('write_after', cell_overwrite_ok=True) 	
    total_rows = 0
    total_cols = 0
    for sheet in sheet_list:
        for rows in range(3, sheet.nrows):
            rows_content = sheet.row_values(rows) # 获取第rows行内容
            if ( rows_content[0] == '' and  rows_content[1] == '' and  rows_content[2] == '' and  rows_content[3] == '' ):
                continue 
            i = 0
            print(rows_content[sheet.nrows])
            for item in rows_content:
                sheet_w.write(total_rows, i, item) 
                i += 1
            total_rows += 1
    wbk.save('C:\\Users\\keithliu\\Desktop\\demo1.xls')
    print ("Total rows is %d" %  total_rows)
if  __name__ == '__main__':
    (fname_list, fhandle_list) = open_all_files("D:\\attachments\\")
    file_num = len( fname_list) 
    print ("Totally open %d excel files" % file_num)
    sheet_list = open_merge_sheet(fhandle_list)
    merge_sheet( sheet_list )
