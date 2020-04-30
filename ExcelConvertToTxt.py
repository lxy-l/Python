import pandas as pd
import codecs
import os
#将excel转化为txt文件
def exceltotxt(excel_dir):  
        excel = pd.ExcelFile(excel_dir)
        for item in (excel.sheet_names):
            neg=pd.read_excel(excel_dir,item,index_col=0)
            for row2 in neg:
                print(row2)
            path=os.getcwd()+"\\"+os.path.splitext(excel_dir)[0]+"\\"
            isExists=os.path.exists(path)
            if not isExists:
                os.makedirs(path)
            for index, row in neg.iterrows():
                with codecs.open(path+str(index)+".txt", 'w', 'gb18030') as f:
                    f.write(row.to_string())   
exceltotxt('Excel.xls')
