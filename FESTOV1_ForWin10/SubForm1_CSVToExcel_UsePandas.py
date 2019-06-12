from tkinter import *
from tkinter import filedialog
import tkinter.messagebox
import os
import csv
from xlwt import *
import time
import pandas as pd
import xlrd
import xlwt

'''
   使用pandas 转换为Excel
   1.先用Pandas读取CSV文件，然后生成xlsx文件
   2.再用xlrd读取第1步生成的xlsx文件，使用xlwt转化生成xls文件，去掉xlsx文件第1列的序号

   说明：
   1.之所以使用Pandas读取CSV，生成Excel文件，是因为使用csv.reader(sFile)，无法解决MainDes里有逗号，导致读取到的数据乱列
   2.之所以再用xlrd读取一次，用xlwt生成文件，是因为使用Pandas生成的Excel里第1列是序号，而且是xlsx格式，后面步骤需要用到的是xls文件格式
'''

def main():
    def selectExcelfile():
        sfname = filedialog.askopenfilename(title='Please Select CSV File', filetypes=[('CSV', '*.csv'), ('All Files', '*')])
        #print(sfname)
        txt_CsvFile.insert(INSERT,sfname)
        root.wm_attributes('-topmost',1)


    def closeThisWindow():
        root.destroy()

    def doProcess():
        root.wm_attributes('-topmost',0)
        tkinter.messagebox.showinfo('提示信息','点OK按钮开始处理CSV文件......')
        root.wm_attributes('-topmost',1)

        sCsvFile=txt_CsvFile.get()
     
        csv = pd.read_csv(sCsvFile, encoding='utf-8')  
       
        strTime=time.strftime('%Y%m%d_%H%M%S',time.localtime(time.time())) 
        sPath=os.getcwd()+"\\Result\\"
        sReportName="1.ClosedTickets_"+strTime+".xlsx"            

        csv.to_excel(sPath+sReportName,sheet_name='Sheet1')


        #把生成的.xlsx文件格式转化为.xls文件格式,同时去掉因Pandas生成的第1列的序号
        sSourceFile = sPath+sReportName
        sTargetFile = sPath + "1.ClosedTickets_"+strTime+".xls" 
        ChangeXlsxToXls(sSourceFile,sTargetFile)

        RemoveMidFile(sSourceFile)
        
        
        
        print("Process is over.")
        root.wm_attributes('-topmost',0)
        tkinter.messagebox.showinfo('提示信息','已经转换生成Excel文件。\n '+sTargetFile)
        root.wm_attributes('-topmost',1)

    def csv_to_xlsx_pd():
        csv = pd.read_csv('1.csv', encoding='utf-8')
        csv.to_excel('1.xlsx', sheet_name='data')

    def ChangeXlsxToXls(sSourceFile,sTargetFile):
        wb1 = xlrd.open_workbook(filename=sSourceFile)
        sheet1 = wb1.sheet_by_index(0)   
        nrows1 = sheet1.nrows
        ncols2 = sheet1.ncols

        wb2 = xlwt.Workbook()
        ws = wb2.add_sheet('Sheet1',cell_overwrite_ok=True)

        for iRow in range(nrows1):
            for iCol in range(1,ncols2):                                     #第1列是使用Pandas生成的Excel,自动生成序号，去掉
                ws.write(iRow,iCol-1,sheet1.cell_value(iRow,iCol))
        wb2.save(sTargetFile)
        
    def RemoveMidFile(sTempFile):
        if os.path.exists(sTempFile):
            os.remove(sTempFile)
    
    
    #初始化
    root=Tk()

    #设置窗体标题
    root.title('CSV file to Excel ---Step 1')

    #设置窗口大小和位置
    root.geometry('660x300+430+220')


    label1=Label(root,text='Select CSV File:')
    txt_CsvFile=Entry(root,bg='white',width=70)
    btn_Browse=Button(root,text='Browse',width=8,command=selectExcelfile)


    btn_Do=Button(root,text='Process',width=8,command=doProcess)
    btn_Exit=Button(root,text='Exit',width=8,command=closeThisWindow)

    label1.pack()
    txt_CsvFile.pack()
    btn_Browse.pack()



    btn_Do.pack()
    btn_Exit.pack()
    

    label1.place(x=30,y=30)
    txt_CsvFile.place(x=120,y=30)
    btn_Browse.place(x=550,y=26)

    
    btn_Do.place(x=230,y=120)
    btn_Exit.place(x=330,y=120)
 
    root.mainloop() 

 
if __name__=="__main__":
    main()
