import sys
import os
import pyodbc 
import pandas as pd





def main():
    
    
    #   -------------------  Configurations -------------------------------
    #                        ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
    

    
     excelOutput=pd.DataFrame()
     xcelOutput=pd.DataFrame()
   

     outputFile = open('DB_Output.txt', 'a')
     print(".............Process Start ............")

     #RESTORE DATABASE [OKAMI_ZiiPOS_BranchName] FILE = N'OKAMI_ZiiPOS' FROM  DISK = N'C:\ZiiBackup\OKAMI_ZiiPOS_V3.bak' WITH  FILE = 1,  MOVE N'OKAMI_ZiiPOS' TO N'c:\Program Files\Microsoft SQL Server\MSSQL10_50.SQLEXPRESS2008R2\MSSQL\DATA\OKAMI_ZiiPOS_BranchName.mdf',  MOVE N'OKAMI_ZiiPOS_log' TO N'c:\Program Files\Microsoft SQL Server\MSSQL10_50.SQLEXPRESS2008R2\MSSQL\DATA\OKAMI_ZiiPOS_BranchName_0.ldf',  NOUNLOAD,  STATS = 10 GO


     dblist = pd.read_excel('dblist_full.xlsx', index_col=None,dtype = str)
     for y in range(len(dblist)):
          print(y ," / ",len(dblist))
          BranchName = dblist.iloc[y]["BranchName"]

          infoQuery = "RESTORE DATABASE [OKAMI_ZiiPOS_"+BranchName+"] FILE = N'OKAMI_ZiiPOS' FROM  DISK =  N'C:\ZiiBackup\OKAMI_ZiiPOS_V3.bak' WITH  FILE = 1,  MOVE N'OKAMI_ZiiPOS' TO N'c:\Program Files\Microsoft SQL Server\MSSQL10_50.SQLEXPRESS2008R2\MSSQL\DATA\OKAMI_ZiiPOS_"+BranchName+".mdf',  MOVE N'OKAMI_ZiiPOS_log' TO N'c:\Program Files\Microsoft SQL Server\MSSQL10_50.SQLEXPRESS2008R2\MSSQL\DATA\OKAMI_ZiiPOS_"+BranchName+"_0.ldf',  NOUNLOAD,  STATS = 10"
          outputFile.write(infoQuery)
          outputFile.write('\n')
          outputFile.write('GO')
          outputFile.write('\n')
          outputFile.write('\n')

          
        

    




main()

