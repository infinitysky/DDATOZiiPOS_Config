import sys
import os
import pyodbc 
import pandas as pd
import numpy as np
import signal




def main():
    
    
    #   -------------------  Configurations -------------------------------
    #                        ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
    server = '192.168.20.232,9899'
    #database = '19Okami_Chelsea_Heights'
    username = 'ZiiPos'
    password = 'ZiiPos884568'

    #Password Connection
    #PassSQLServerConnection = pyodbc.connect('DRIVER={SQL Server}; SERVER='+server+'; DATABASE='+database+'; UID='+username+'; PWD='+ password)
    
    #MenuItemQuery = "SELECT  ItemCode ,Description1 ,Description2 ,Category ,PrinterPort ,PrinterPort1 ,PrinterPort2 ,PrinterPort3 FROM MenuItem order by ItemCode"
    
    
    #OutputFileName = database + '_outputFile.xlsx';
    
    #                        ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑            
    #   -------------------  Configurations -------------------------------
    
    
    print(".............Process Start ............")
    dblist = pd.read_excel('dblist_full.xlsx', index_col=None,dtype = str)
    
    NewExcelDataFrame=pd.DataFrame()
    for y in range(len(dblist)):
        print(y ," / ",len(dblist))
        
        databaseName = dblist.iloc[y]["dblist"]
        print("Process Database :" + databaseName)
        PassSQLServerConnection = pyodbc.connect('DRIVER={SQL Server}; SERVER='+server+'; DATABASE='+databaseName+'; UID='+username+'; PWD='+ password)
    
        MenuItemQuery = "SELECT CompanyName,Telephone,Fax,ABN,Address FROM Profile"
    
        SqlResult1 = pd.read_sql(MenuItemQuery, PassSQLServerConnection)
        SourceDataFromDB = SqlResult1.astype("string")
        NewExcelDataFrame = NewExcelDataFrame.append(SourceDataFromDB)
        

        PassSQLServerConnection.close()
        print(databaseName + " Completed")
        print("")
        
    
    print(".............All Process completed.............")
    NewExcelDataFrame.to_excel("info.xlsx", index = True, header=True)










main()

