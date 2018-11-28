# -*- coding: utf-8 -*-
"""
Created on Tue Jul  3 11:22:02 2018

@author: yangyanhao
"""

import requests
import xlwt
import win32com.client

conn = win32com.client.Dispatch('ADODB.Connection')
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=I9500.mdb"
conn.Open()
rs = win32com.client.Dispatch('ADODB.Recordset')
rs.Open('Reports',conn,1,3) #Open(sql,conn,1,3) 
rs.MoveLast()



params = {}
def get_Inf(code):
    params['Barcode'] =code
    #r = requests.get("/gtin/SearchGtin", data=params)
    Trans=eval(r.text) 
    return(Trans)
    
def write_Excel(code,i,write_file_name):
    Code=get_Inf(code)
    #write_file_name ='d:\\UserData\\yangyanhao\\Desktop\\记得改名.xls'
    #workbook = xlwt.Workbook(write_file_name)
    #worksheet = workbook.add_sheet('TEST')
    worksheet.write(i, 0, label = Code['gtin'][1:])
    worksheet.write(i, 1, label = Code['productName'])
    worksheet.write(i, 2, label = Code['content'])
    worksheet.write(i, 3, label = Code['partyContactName'])
    #workbook.save(write_file_name)
    return()


#Code=rs.Fields.Item(11).Value
#write_Excel(k,1)

write_file_name ='周五检测.xls'
workbook = xlwt.Workbook(write_file_name)
worksheet = workbook.add_sheet('服务性检测')

i=30
for k in range(i):
    Code=rs.Fields.Item(11).Value
    write_Excel(Code,k,worksheet)
    worksheet.write(k, 4, label =rs.Fields.Item(10).Value)
    rs.MovePrevious()
    
workbook.save(write_file_name)
