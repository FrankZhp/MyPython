#! -*- coding utf-8 -*-
#! @Time  :2019/3/4 9:00
#! Author :Frank Zhang
#! @File  :ProjectMain.py
#! Python Version 3.7
import tkinter
import tkinter.messagebox
import SubForm1_CSVToExcel_UsePandas
import SubForm2_MergeFESTO
import SubForm3_Mapping
import SubForm4_Summary
import SubForm5_MergeInvoice
import SubForm6_AppendInvoiceToSummary
import SubForm7_GetUnitPrices
import SubForm8_GetUnitPrices
import SubForm9_ComputePrice
#import SubForm9_MakeReport
#import SubForm10_MakeReport


def main():
   
    #设置窗体标题
    root.title('Compute Closed Tickets Price')

    #设置窗口大小和位置
    root.geometry('800x500+360+100')

    menu = tkinter.Menu(root)
    root.config(menu=menu)
    submenu1 = tkinter.Menu(menu,tearoff=0)
    submenu1.add_command(label="1.CSV File To Excel",command=openCSVToExcelForm)
    submenu1.add_command(label="2.Merge FESTO Files",command=openMergeFESTOForm)
    submenu1.add_command(label="3.Mapping",command=openMappingForm)
    submenu1.add_command(label="4.Summary",command=openSummaryForm)
    submenu1.add_command(label="5.Merge Invoice",command=openMergeInvoiceForm)
    submenu1.add_command(label="6.Append Invoice to Summary",command=openAppendInvoiceToSummaryForm)
    submenu1.add_command(label="7.Get Unit Prices Style 1",command=openGetUnitPices1Form)
    submenu1.add_command(label="8.Get Unit Prices Style 2",command=openGetUnitPices2Form)
    submenu1.add_command(label="9.Compute Tickets Price" ,command=openComputePriceForm)
    #submenu1.add_command(label="9.Make Report Style 2",command=openMakeReport1Form)
    #submenu1.add_command(label="10.Make Report Style 3",command=openMakeReport2Form)
    submenu1.add_separator()
    submenu1.add_command(label="Exit",command=closeThisWindow)
    menu.add_cascade(label="Work",menu=submenu1)

    submenu2 = tkinter.Menu(menu,tearoff=0)
    submenu2.add_command(label="Help",command=openHelp)
    submenu2.add_command(label="about",command=openAbout)
    menu.add_cascade(label="Help",menu=submenu2)

    root.mainloop()



def openCSVToExcelForm():
    print("Open CSV File To Excel Form ---Step 1")
    SubForm1_CSVToExcel_UsePandas.main()

def openMergeFESTOForm():
    print("open Merge FESTO Form ---Step 2")
    SubForm2_MergeFESTO.main()

def openMappingForm():
    print("open Mapping Form ---Step 3")
    SubForm3_Mapping.main()
    
def openSummaryForm():
    print("open Summary Form ---Step 4")
    SubForm4_Summary.main()
    
def openMergeInvoiceForm():
    print("open Merge Invoice Form ---Step 5")
    SubForm5_MergeInvoice.main()

def openAppendInvoiceToSummaryForm():
    print("open Append Invoice To Summary Form ---Step 6")
    SubForm6_AppendInvoiceToSummary.main()

def openGetUnitPices1Form():
    print("open Get Unit Prices Style1 Form ---Step 7")
    SubForm7_GetUnitPrices.main()

def openGetUnitPices2Form():
    print("open Get Unit Prices Style2 Form ---Step 8")
    SubForm8_GetUnitPrices.main()

def openComputePriceForm():
    print("open Compute Price Form ---Step 9")
    SubForm9_ComputePrice.main()

'''
2019-05-08目前用不到
def openMakeReport1Form():
    print("open Make Report 1 Form")
    SubForm9_MakeReport.main()

def openMakeReport2Form():
    print("open Make Report 2 Form")
    SubForm10_MakeReport.main()
'''
  
def closeThisWindow():
    if  tkinter.messagebox.askokcancel('提示', '确定要退出吗？')==True:
        root.destroy()

def openHelp():
     tkinter.messagebox.showinfo("欢迎", "欢迎使用Compute Tickets Price Tool！\nBy Frank Zhang")
  

def openAbout():
    print("openAbout")

if __name__=="__main__":
    root = tkinter.Tk()
    main()
                    
                    
                    
