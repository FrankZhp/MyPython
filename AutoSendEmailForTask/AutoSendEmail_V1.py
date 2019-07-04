# -*- coding: utf-8 -*-

"""
Module 根据定义的Task，发送邮件提醒，避免遗漏处理
"""

import os
import xlrd
import schedule
import time
import win32com.client as win32
from datetime import datetime,date


dictWeek={0:"Monday",1:"TuesDay",2:"Wednesday",3:"Thursday",4:"Friday",5:"Saturday",6:"Sunday"}

def main():
    sPath = os.getcwd()
    sFile = "TaskList.xlsx"
    sExcelFile = sPath +"\\" + sFile
    
    wb = xlrd.open_workbook(filename=sExcelFile)
      
    sheet1 = wb.sheet_by_index(0)   
    nrows1 = sheet1.nrows
    
    #注意weekday() 返回的是0-6是星期一到星期日
    sWeekday = dictWeek.get(datetime.now().weekday())
    sNow = datetime.now() 

    iDay = sNow.day
    sToday= formatDay(sNow,"yyyy-mm-dd")

    for iRow in range(1,nrows1):
        sCheck = sheet1.cell(iRow,3).value
        if sCheck != "Y":
            sFrequency = sheet1.cell(iRow,0).value
            sItem = sheet1.cell(iRow,1).value
            if sFrequency == "Week":
                if sItem == sWeekday:
                    sTask = sheet1.cell(iRow,2).value
                    sendEmail(sTask)
            elif sFrequency == "Day":
                #sItem为字符串类型，只有当sItem的长度小于10的时候，再进一步检查转换
                if len(sItem) < 10:
                    if isVaildDate(sItem):
                        t1 = time.strptime(sItem, "%Y-%m-%d")
                        sItem = changeStrToDate(t1,"yyyy-mm-dd") 
                if sItem == sToday:
                    sTask = sheet1.cell(iRow,2).value
                    sendEmail(sTask)
            elif sFrequency == "Month":
                if int(sItem) == int(iDay):
                    sTask = sheet1.cell(iRow,2).value
                    sendEmail(sTask)
                       
def formatDay(sDay,sFormat):
    sYear = str(sDay.year)
    sMonth = str(sDay.month)
    sDay = str(sDay.day)

    if sFormat == "yyyy-mm-dd":
        sFormatDay = sYear +"-" +sMonth.zfill(2)+"-" +sDay.zfill(2)
    elif sFormatStyle == "yyyy/mm/dd":
        sFormatDay = sYear +"/" +sMonth.zfill(2)+"/" +sDay.zfill(2)
    else:
        sFormatDay = sYear+"-" + sMonth + "-" + sDay
        
    return sFormatDay

"""
功能：判断是否为日期
"""
def isVaildDate(sDate):
    try:
        if ":" in sDate:
            time.strptime(sDate, "%Y-%m-%d %H:%M:%S")
        else:
            time.strptime(sDate, "%Y-%m-%d")
        return True
    except:
        return False

"""
   功能：把字符串格式的日期转换为格式化的日期，如把2019-7-1转换为2019-07-01
"""
def changeStrToDate(sDate,sFormat):
    sYear = str(sDate.tm_year)
    sMonth = str(sDate.tm_mon)
    sDay = str(sDate.tm_mday)

    if sFormat == "yyyy-mm-dd":
        sFormatDay = sYear +"-" +sMonth.zfill(2)+"-" +sDay.zfill(2)
    elif sFormatStyle == "yyyy/mm/dd":
        sFormatDay = sYear +"/" +sMonth.zfill(2)+"/" +sDay.zfill(2)
    else:
        sFormatDay = sYear+"-" + sMonth + "-" + sDay
        
    return sFormatDay
    
def sendEmail(sTask):
    try:
	#读取config.txt，获得发送的目标邮箱账号
        sConfigFile="config.txt" 
            
        f=open(sConfigFile,'r')
        try:
            file_Context=f.read()
        except:
            return False
        finally:
            if f:
                f.close()
		
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        receivers = [file_Context]
        mail.To = receivers[0]
        mail.Subject ='这是一封提醒邮件.'
        mail.Body="邮件提醒：  \r\n    请注意处理任务作业，如已处理可忽略此封邮件。\r\n   任务内容：" + sTask + " \r\n     (此邮件由系统自动发送)"
        #mail.Attachments.Add('C:\\Users\enegc\\OneDrive - Bayer\\Personal Data\\'+sFileName+'.xlsx')
        mail.Send()
        return True
    except exceptions as e:
        return False

if __name__ == "__main__":
    main()
 
