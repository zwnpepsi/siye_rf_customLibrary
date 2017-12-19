import  xlrd
import tkinter.filedialog
import xlsxwriter


class DoExcel(object):
    '''
    A test library for excel
    '''
    ROBOT_LIBRARY_VERSION = 1.0
    ROBOT_LIBRARY_SCOPE='GLOBAL'

    def __init__(self):
        pass

    def read_excel(self):
        workbook=xlrd.open_workbook(tkinter.filedialog.askopenfilename(title="打开指定的excel文档"))
        sheet_obj = workbook.sheets()[0] #获取sheet1   #根据索引获取excel里面的表单对象
        col1=sheet_obj.col_values(0)     #取第1列的值
        row1=sheet_obj.row_values(0)
#         for i in range(sheet_obj.nrows):
#             print(col1[i])
        for i in col1:
            print(i)
        
    
    def write_excel(self,i,j,value):
        workbook = xlsxwriter.Workbook(tkinter.filedialog.asksaveasfilename(filetypes=[('xls','.xls'),('xlsx','.xlsx'),('All file','*')],title="保存指定的excel文档",defaultextension='.xlsx'))
        worksheet=workbook.add_worksheet('test')

        worksheet.write(int(i),int(j),value)   #写数据
        workbook.close() 
        

    
