import xlrd
import csv
import os
import codecs
import tkinter
import tkinter.filedialog as filedialog
import tkinter.messagebox as tkMessageBox


def GetExcelArray():
    ls = []
    f = []
    for root, dirs, files in os.walk(os.getcwd()):
        for file in files:
            if os.path.splitext(file)[1] == ".xls" or os.path.splitext(file)[1] == ".xlsx":
                ls.append(os.path.join(root, file))
                f.append(file)
    return f


def xlsx_to_csvList(pathArray):
    for excel_dir in pathArray:
        workbook = xlrd.open_workbook(excel_dir)
        for index in range(0, len(workbook.sheet_names())):
            table = workbook.sheet_by_index(index)
            path = os.getcwd()+"\\" + \
                os.path.splitext(excel_dir)[0]+"\\"+table.name+"\\"
            isExists = os.path.exists(path)
            if not isExists:
                os.makedirs(path)
            for row_num in range(table.nrows):
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

def xlsx_to_csvOnlyOne(excel_dir):
    workbook = xlrd.open_workbook(excel_dir)
    for index in range(0,len(workbook.sheet_names())):
        table=workbook.sheet_by_index(index)
        # father_path=os.path.abspath(os.path.dirname(excel_dir)+os.path.sep+".")
        # path=os.getcwd()+"\\"+os.path.splitext(excel_dir)[0]+"\\"+table.name+"\\"
        path=os.path.abspath(os.path.dirname(excel_dir)+os.path.sep+".")+"\\BuglistCSV\\"+table.name+"\\"
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

def ShowMessage():
    tkMessageBox.showinfo("提示", "转换完成！")
    tkMessageBox.askokcancel('askokcancel', 'hello')
    tkMessageBox.askretrycancel('askretrycancel', 'hello')


def QuestionConvert(path):
    # tkMessageBox.askquestion('询问', '是否开始转换？')
    isTrue = tkMessageBox.askyesno('询问', '是否开始转换？')
    if(isTrue):
        # tkMessageBox.showinfo("提示", path)
        xlsx_to_csvOnlyOne(path)
        # files=GetExcelArray()
        # if(files==[]):
        #     tkMessageBox.showwarning('提示', '未找到Excel文件！')
        # else:
        #     xlsx_to_csv(files)
    # else:
    #     tkMessageBox.showwarning('提示', '取消操作！')
        # tkMessageBox.showerror('警告', '转换失败！')


def OptionFile():
    filename = filedialog.askopenfilename()
    if filename != '':
        if os.path.splitext(filename)[1] == ".xls" or os.path.splitext(filename)[1] == ".xlsx":
            QuestionConvert(filename)
        else:
            tkMessageBox.showerror('警告', '请选择Excel文件！')
        #  lb.config(text='Path:'+filename)
        
    # else:
    #     lb.config(text='未选择任何文件')


# root = tkinter.Tk()
# root.title('ExcelToCSV')
# root.geometry('300x150')

# # lb = tkinter.Label(root,text='')
# btn = tkinter.Button(root, text='选择文件', command=OptionFile)
# # btn2=tkinter.Button(root,text='开始转换',command=lambda: QuestionConvert(lb))

# # lb.pack()
# btn.pack()
# # btn2.pack()
# root.mainloop()
# xlsx_to_csvList(GetExcelArray())
tkinter.Tk().withdraw();
OptionFile()