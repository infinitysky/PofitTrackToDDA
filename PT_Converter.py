import tkinter as tk
import tkinter.font as tkFont
from tkinter import messagebox
import sys
import os
import pyodbc 
import pandas as pd
import numpy as np
import xlsxwriter
import xlrd
import wget


def closesystem():
    sys.exit()

def convertToDDAExcel(SourceData,DDATemplate):
    DDAExcelDataFrame = pd.DataFrame()

    TemplateData=DDATemplate.iloc[0]
    TemplateData["XeroTaxCode"]="GST on income"
    TemplateData["WholesalePrice1(Inc GST)"]=0
    TemplateData["WholesalePrice2(Inc GST)"]=0


    z=0
    
    for z in range(len(SourceData)):
        print(z ," / ",len(SourceData) ," (2)" )
        if pd.notna(SourceData.iloc[z]["Product Code"]):
            TemplateData=DDATemplate.iloc[0]
            TemplateData["ProductCode(15)"] = SourceData.iloc[z]["Product Code"]
            TemplateData["Description1(100)"] = SourceData.iloc[z]["Product Description"]
            TemplateData["Description2(100)"] = SourceData.iloc[z]["Caption"]
            
                
            TemplateData["Category(25)"] = SourceData.iloc[z]["Group Name"]
            
            TemplateData["SalesPrice1(Inc GST)"] = SourceData.iloc[z]["Sell Price 1"]
            TemplateData["SalesPrice2(Inc GST)"] = SourceData.iloc[z]["Sell Price 2"]
            TemplateData["SalesPrice3(Inc GST)"] = SourceData.iloc[z]["Sell Price 3"]
            TemplateData["SalesPrice4(Inc GST)"] = SourceData.iloc[z]["Sell Price 4"]
            TemplateData["SalesPrice5(Inc GST)"] = SourceData.iloc[z]["Sell Price 5"]
            TemplateData["SalesPrice6(Inc GST)"] = SourceData.iloc[z]["Sell Price 6"]
       
            TemplateData["Barcode1(30)"] = SourceData.iloc[z]["barcode1"]
            TemplateData["Barcode2(30)"] = SourceData.iloc[z]["barcode2"]
            TemplateData["Barcode3(30)"] = SourceData.iloc[z]["barcode3"]
            TemplateData["Barcode4(30)"] = SourceData.iloc[z]["barcode4"]
            TemplateData["Barcode5(30)"] = SourceData.iloc[z]["barcode5"]
            TemplateData["Barcode6(30)"] = SourceData.iloc[z]["barcode6"]

          
            TemplateData["DeductStockQty1"] = SourceData.iloc[z]["Units Per Plu1"]
            TemplateData["DeductStockQty2"] = SourceData.iloc[z]["Units Per Plu2"]
            TemplateData["DeductStockQty3"] = SourceData.iloc[z]["Units Per Plu3"]
            TemplateData["DeductStockQty4"] = SourceData.iloc[z]["Units Per Plu4"]
            TemplateData["DeductStockQty5"] = SourceData.iloc[z]["Units Per Plu5"]
            TemplateData["DeductStockQty6"] = SourceData.iloc[z]["Units Per Plu6"]
           
            TemplateData["Default Supplier Code (15)"] = SourceData.iloc[z]["Supplier Code"]
            TemplateData["GSTRate"] = SourceData.iloc[z]["GST Rate"]

            TemplateData["LastOrderPrice(Ex GST)"] = SourceData.iloc[z]["Last Buy Cost"]



        



            TemplateData = TemplateData.to_frame()
            TemplateData = TemplateData.transpose()
                

            DDAExcelDataFrame = pd.concat([DDAExcelDataFrame, TemplateData],ignore_index=True)



    return DDAExcelDataFrame




def processBarcodes(productList,prodproductCodeList):
    finalResult = pd.DataFrame()
    finalResult["Product Code"]=""
    finalResult["GST Rate"]=""
    finalResult["Product Description"]=""
    finalResult["Supplier Code"]=""
    finalResult["Supplier Name"]=""
    
    finalResult["Caption"]=""
    
    finalResult["Group Name"]=""
    finalResult["Group Name"]=""
    finalResult["Department Number"]=""
    
    finalResult["Department Name"]=""
    
    finalResult["Sell Price 1"]=""
    finalResult["Sell Price 2"]=""
    finalResult["Sell Price 3"]=""
    finalResult["Sell Price 4"]=""
    finalResult["Sell Price 5"]=""
    finalResult["Sell Price 6"]=""

    finalResult["ItemDescription1"]=""
    finalResult["ItemDescription2"]=""
    finalResult["ItemDescription3"]=""
    finalResult["ItemDescription4"]=""
    finalResult["ItemDescription5"]=""
    finalResult["ItemDescription6"]=""

    
    finalResult["Units Per Plu1"]=""
    finalResult["Units Per Plu2"]=""
    finalResult["Units Per Plu3"]=""
    finalResult["Units Per Plu4"]=""
    finalResult["Units Per Plu5"]=""
    finalResult["Units Per Plu6"]=""


    finalResult["barcode1"]=""
    finalResult["barcode2"]=""
    finalResult["barcode3"]=""
    finalResult["barcode4"]=""
    finalResult["barcode5"]=""
    finalResult["barcode6"]=""
    finalResult["Last Buy Cost"]=""
    
    z=0
    lenth=len(prodproductCodeList)
    print(lenth)
   

    for z in range(len(prodproductCodeList)):
        print(z ," / ",lenth, " (1)" )
        TemplateData = prodproductCodeList.iloc[z]
        
        #print(TemplateData)
        T1=productList[productList["Product Code"] == TemplateData['Product Code']]
        #print(T1)
        barCodeLength = len(T1)
        
        #print(barCodeLength)
        if barCodeLength > 0:
            y=0
            for y in range (barCodeLength):
                if T1.iloc[y]["Plu"] !=TemplateData["Product Code"]:
                    name="barcode"+ str(y+1)
                    TemplateData[name]=T1.iloc[y]["Plu"]

                    salesPriceName="Sell Price "+ str(y+1)
                    TemplateData[salesPriceName]=T1.iloc[y]["Sell Price 1"]

                    itemName="ItemDescription"+str(y+1)
                    TemplateData[itemName]=T1.iloc[y]["Caption"]

                    unitPerPluName="Units Per Plu"+str(y+1)
                    TemplateData[unitPerPluName]=T1.iloc[y]["Units Per Plu"]

                elif T1.iloc[y]["Plu"] ==TemplateData["Product Code"]:

                    name="barcode"+ str(y+1)
                    TemplateData[name]=str(T1.iloc[y]["Plu"])+"-1"

                    salesPriceName="Sell Price "+ str(y+1)
                    TemplateData[salesPriceName]=T1.iloc[y]["Sell Price 1"]

                    itemName="ItemDescription"+str(y+1)
                    TemplateData[itemName]=T1.iloc[y]["Caption"]

                    unitPerPluName="Units Per Plu"+str(y+1)
                    TemplateData[unitPerPluName]=T1.iloc[y]["Units Per Plu"]
                   
        
    
        TemplateData["Product Code"]=T1.iloc[0]["Product Code"]
        TemplateData["Product Description"]=T1.iloc[0]["Product Description"]

        TemplateData["GST Rate"]=T1.iloc[0]["GST Rate"]

        TemplateData["Supplier Code"]=T1.iloc[0]["Supplier Code"]
        TemplateData["Supplier Name"]=T1.iloc[0]["Supplier Name"]
        
        TemplateData["Caption"]=T1.iloc[0]["Caption"]
        
        TemplateData["Group Name"]=T1.iloc[0]["Group Name"]
        TemplateData["Group Number"]=T1.iloc[0]["Group Number"]
        TemplateData["Department Name"]=T1.iloc[0]["Department Name"]
        TemplateData["Department Number"]=T1.iloc[0]["Department Number"]
        TemplateData["Last Buy Cost"]=T1.iloc[0]["Last Buy Cost"]
    

        TemplateData = TemplateData.to_frame()
        TemplateData = TemplateData.transpose()
        finalResult = pd.concat([finalResult, TemplateData],ignore_index=True)

    # for z in range(len(productList)):
    #     print(z ," / ",lenth, " (1)" )
    #     TemplateData = productList.iloc[z]
        
        
    #     T1= pd.DataFrame()
        
    #     T1=prodproductCodeList[prodproductCodeList["stock_id"] == TemplateData['stock_id']]
  
    #     barCodeLength = len(T1)
        
    #     #print(barCodeLength)
    #     if barCodeLength > 1:
    #         y=0
    #         for y in range (barCodeLength):
    #             if T1.iloc[y]["barcode"] !=TemplateData["barcode"]:
    #                 name="barcode"+ str(y+1)
    #                 TemplateData[name]=T1.iloc[y]["barcode"]
            

    #     TemplateData["stock_id"]= int(TemplateData["stock_id"])
    #     TemplateData = TemplateData.to_frame()
    #     TemplateData = TemplateData.transpose()
    #     finalResult = pd.concat([finalResult, TemplateData],ignore_index=True)
           

    return finalResult


         


def processProductWithBarCode(connect_string):
    PassSQLServerConnection = pyodbc.connect(connect_string)

  

    productCodeListQuery = "SELECT [ProductList].[Product Code] FROM [ProductList] group by [Product Code] "
    productListQuery = "SELECT [ProductList].* FROM [ProductList] order by [Product Code]"

   

    prodproductCodeList = pd.read_sql_query(productCodeListQuery, PassSQLServerConnection)
    
    productListQuery =  pd.read_sql_query(productListQuery, PassSQLServerConnection)

    # productListQuery=productListQuery.astype(str)
    # prodproductCodeList=prodproductCodeList.astype(str)


    result=processBarcodes(productListQuery,prodproductCodeList)
    print("total rows: ")
    print(len(result))

 

    print("Export to new excel")
    result.to_excel(r'export_dataframe.xlsx', index = True, header=True,engine='xlsxwriter')
    print("Stage 1 Process completed")

    print("Start convert to DDA")
    DDAExcel = pd.DataFrame()
    try:
        downloadURL="https://download.ziicloud.com/programs/ziiposclassic/ItemImportFormat.xls"
        local_file="ItemImportFormat.xls"
        wget.download(downloadURL, local_file)
    except wget.Error as ex:
        print("Download Files error")
        

    DDADataTemplete = pd.read_excel('ItemImportFormat.xls', index_col=None,dtype = str)
    DDAExcel=DDADataTemplete.astype(str)
    DDAExcel =DDADataTemplete

    DDAExcel = convertToDDAExcel(result,DDAExcel)

    DDAExcel.to_excel(r'outPut.xlsx', index = False, header=True,engine='xlsxwriter')
    messagebox.showinfo(title="Process Completed",message="Data Process Completed")




    
def ConnectionTest(connect_string):
    connectionTestResult = 0
   
    PassSQLServerConnection = pyodbc.connect(connect_string)

    print(connect_string)
    try:
        PassSQLServerConnection = pyodbc.connect(connect_string)
        print("{c} is working".format(c=connect_string))
        PassSQLServerConnection.close()
        connectionTestResult = 1
    except pyodbc.Error as ex:
        #print("{c} is not working".format(c=connect_string))
        messagebox.showerror(title="Error", message="{c} is not working")

    return connectionTestResult


def inforProcess(DBSource,DBUsername,DBPassword,DBName):
    connectionTestResult=0
    connect_string = 'DRIVER={SQL Server}; SERVER='+DBSource+'; DATABASE='+DBName+'; UID='+DBUsername+'; PWD='+ DBPassword
    if DBName=="":
        messagebox.showerror(title="Error", message="DB Name Field is Empty!!")
        connectionTestResult = 0
    else:
        connectionTestResult=ConnectionTest(connect_string)

    if connectionTestResult==1:
        print("next")
        processProductWithBarCode(connect_string)

    else:
        print("error")







class App:
    def __init__(self, root):
        #setting title
        root.title("PT Converter")
        #setting window size
        width=600
        height=500
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        GLabel_DB_Source=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        GLabel_DB_Source["font"] = ft
        GLabel_DB_Source["fg"] = "#333333"
        GLabel_DB_Source["justify"] = "left"
        GLabel_DB_Source["text"] = "DB Connection"
        GLabel_DB_Source.place(x=50,y=90,width=90,height=30)

        DBSource_Box=tk.Entry(root)
        DBSource_Box["borderwidth"] = "1px"
        ft = tkFont.Font(family='Times',size=10)
        DBSource_Box["font"] = ft
        DBSource_Box["fg"] = "#333333"
        DBSource_Box["justify"] = "left"
        DBSource_Box.insert(0,'localhost\\sqlexpress2008r2')
        #DBSource_Box["text"] = "localhost\sqlexpress2008r2"
        
        DBSource_Box.place(x=190,y=90,width=275,height=30)

        DB_UserName_Label=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        DB_UserName_Label["font"] = ft
        DB_UserName_Label["fg"] = "#333333"
        DB_UserName_Label["justify"] = "left"
        DB_UserName_Label["text"] = "User Name"
        DB_UserName_Label.place(x=50,y=140,width=90,height=30)

        DB_UserName_Box=tk.Entry(root)
        DB_UserName_Box["borderwidth"] = "1px"
        ft = tkFont.Font(family='Times',size=10)
        DB_UserName_Box["font"] = ft
        DB_UserName_Box["fg"] = "#333333"
        DB_UserName_Box["justify"] = "left"
        #DB_UserName_Box["text"] = "sa"
        DB_UserName_Box.insert(0,'sa')
        DB_UserName_Box.place(x=190,y=140,width=275,height=30)

        DB_Password_Label=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        DB_Password_Label["font"] = ft
        DB_Password_Label["fg"] = "#333333"
        DB_Password_Label["justify"] = "left"
        DB_Password_Label["text"] = "Password"
        DB_Password_Label.place(x=50,y=200,width=90,height=25)

        DB_Password_Box=tk.Entry(root)
        DB_Password_Box["borderwidth"] = "1px"
        
        ft = tkFont.Font(family='Times',size=10)
        DB_Password_Box["font"] = ft
        DB_Password_Box["fg"] = "#333333"
        DB_Password_Box["justify"] = "left"
        #DB_Password_Box["text"] = "0000"
        DB_Password_Box.insert(0,'0000')
        DB_Password_Box.place(x=190,y=200,width=275,height=30)
        DB_Password_Box["show"] = "*"

        
        
        DB_Name_Label=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        DB_Name_Label["font"] = ft
        DB_Name_Label["fg"] = "#333333"
        DB_Name_Label["justify"] = "left"
        DB_Name_Label["text"] = "DB Name"
        DB_Name_Label.place(x=50,y=260,width=90,height=25)

        DB_Name_Box=tk.Entry(root)
        DB_Name_Box["borderwidth"] = "1px"
        ft = tkFont.Font(family='Times',size=10)
        DB_Name_Box["font"] = ft
        DB_Name_Box["fg"] = "#333333"
        DB_Name_Box["justify"] = "left"    
        DB_Name_Box.place(x=190,y=260,width=275,height=30)
        DB_Name_Box.insert(0,'PTDB')
 









         #-----------------Functions---------------------------------
        def getDBSource():
            result=DBSource_Box.get()
            return result
           
            
        def getDBUsername():
            result=DB_UserName_Box.get()
            return result
      
        def getDBPassword():
            result=DB_Password_Box.get()
            return result
        
        def getDBName():
            result=DB_Name_Box.get()
            return result
      
        
        def StartConversionProcess():
            DBSource=getDBSource()
            username=getDBUsername()
            password=getDBPassword()
            databaseName=getDBName()
            #inforProcess(DBSource,DBUsername,DBPassword,DBName):
            inforProcess(DBSource,username,password,databaseName)
            

        def testDBSource():
            DBSource=getDBSource()
            username=getDBUsername()
            password=getDBPassword()
            databaseName=getDBName()
            connect_string = 'DRIVER={SQL Server}; SERVER='+DBSource+'; DATABASE='+databaseName+'; UID='+username+'; PWD='+ password
            PassSQLServerConnection = pyodbc.connect(connect_string)
            if databaseName=="":
                messagebox.showerror(title="Error", message="DB Name Field is Empty!!")
            else:
                try:
                    PassSQLServerConnection = pyodbc.connect(connect_string)
                    print("{c} is working".format(c=connect_string))
                    PassSQLServerConnection.close()
                except pyodbc.Error as ex:
               
                    print("{c} is not working".format(c=connect_string))
                    messagebox.showerror(title="Error", message="{c} is not working")
          
            




            
            


            
            
        
        
        

            















        



            
#--------------Button Actions-------------------------
        Star_Button=tk.Button(root)
        Star_Button["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        Star_Button["font"] = ft
        Star_Button["fg"] = "#000000"
        Star_Button["justify"] = "center"
        Star_Button["text"] = "Start"
        Star_Button.place(x=70,y=390,width=90,height=45)
        Star_Button["command"] = StartConversionProcess

        TEST_Button=tk.Button(root)
        TEST_Button["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        TEST_Button["font"] = ft
        TEST_Button["fg"] = "#000000"
        TEST_Button["justify"] = "center"
        TEST_Button["text"] = "Test DB"
        TEST_Button.place(x=250,y=390,width=90,height=45)
        TEST_Button["command"] = testDBSource

        Close_Button=tk.Button(root)
        Close_Button["bg"] = "#f0f0f0"
        ft = tkFont.Font(family='Times',size=10)
        Close_Button["font"] = ft
        Close_Button["fg"] = "#000000"
        Close_Button["justify"] = "center"
        Close_Button["text"] = "Close"
        Close_Button.place(x=420,y=390,width=90,height=45)
        Close_Button["command"] = closesystem
       
        




#----------------Not in use--------------------------------
    def Star_Button_command(self):
        print("Star_Button_command")
    def TEST_Button_command(self):
        print("command")
    def Close_Button_command(self):
        print("Exit")
        exit()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
