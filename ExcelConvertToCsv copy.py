import xlrd
import csv
import os
import codecs
import tkinter
import tkinter.filedialog as filedialog
import tkinter.messagebox as tkMessageBox

def ConvertToCSV():
    filename = filedialog.askopenfilename()
    if filename != '':
        if os.path.splitext(filename)[1] == ".xls" or os.path.splitext(filename)[1] == ".xlsx":
            if(tkMessageBox.askyesno('询问', '是否开始转换？')):
                workbook = xlrd.open_workbook(filename)
                for index in range(0,len(workbook.sheet_names())):
                    table=workbook.sheet_by_index(index)
                    path=os.path.abspath(os.path.dirname(filename)+os.path.sep+".")+"\\BuglistCSV\\"+table.name+"\\"
                    isExists=os.path.exists(path)
                    if not isExists:
                        os.makedirs(path)
                    for row_num in range(1,table.nrows):
                        row_value = table.row_values(row_num)
                        headers = table.row_values(0)
                        product = table.cell_value(row_num,1)
                        title = table.cell_value(row_num,3)
                        bugnum = table.cell_value(row_num,2)
                        name_csv=str(product)+'_'+str(title)+'_'+str(bugnum)
                        with codecs.open(path+name_csv+'.csv', 'w', encoding='gb18030') as f:
                            write = csv.writer(f)
                            write.writerow(headers)
                            write.writerow(row_value)  
                tkMessageBox.showinfo("提示", "转换完成！")
        else:
            tkMessageBox.showerror('警告', '请选择Excel文件！')
tkinter.Tk().withdraw();
ConvertToCSV()