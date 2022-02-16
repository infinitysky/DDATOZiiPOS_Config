from doctest import master
from operator import indexOf
from pickle import TRUE
from queue import Empty
import sys
import os
import pyodbc 
import pandas as pd
import numpy as np
import signal




def main():
    
    
    #   -------------------  Configurations -------------------------------
    #                        ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
    SourceServer = '127.0.0.1,9899'
    targetServer = '192.168.20.242,9899'
    username = 'sa'
    password = '0000'
    databaseName = master
    RestoreFileName = "OKAMI_ZiiPOS"
    #                        ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑            
    #   -------------------  Configurations -------------------------------
    
    



    print(".............Process Start ............")
    dblist = pd.read_excel('dblist_fix.xlsx', index_col=None,dtype = str)
    for y in range(len(dblist)):
        print(y ," / ",len(dblist))
        BranchName = dblist.iloc[y]["BranchName"]
        SourceDBName = dblist.iloc[y]["dblist"]
        TargetDBName = "OKAMI_ZiiPOS_" + BranchName

        print(".............Processing "+ BranchName +"............")

        SourceSQLServerConnection = pyodbc.connect('DRIVER={SQL Server}; SERVER='+SourceServer+'; DATABASE='+SourceDBName+'; UID='+username+'; PWD='+ password)
        TargetSQLServerConnection = pyodbc.connect('DRIVER={SQL Server}; SERVER='+targetServer+'; DATABASE='+TargetDBName+'; UID='+username+'; PWD='+ password)
        TargetSQLServerConnection.autocommit = TRUE
        TargetSQLServerCursor = TargetSQLServerConnection.cursor()
        
       
        

        #---------------- 1. Branch Profile --------------------------------------

        print("Start Process Branch Profil")
        infoQuery = "Select CompanyName, Telephone,ABN,Address,JobListFormatForPrinter1,JobListFormatForPrinter2,JobListFormatForPrinter3,JobListFormatForPrinter4,JobListFormatForPrinter5,JobListFormatForPrinter6,JobListFormatForPrinter7,JobListFormatForPrinter8,JobListFormatForPrinter9,JobListFormatForPrinter10,JobListFormatForPrinter11,JobListFormatForPrinter12 From Profile;"
    
        SqlResult1 = pd.read_sql(infoQuery, SourceSQLServerConnection)
        SourceDataFromDB = SqlResult1.astype("string")

           
        StoreCompanyName=SourceDataFromDB.iloc[0]["CompanyName"]
        StoreTelephone = SourceDataFromDB.iloc[0]["Telephone"]
        StoreABN = SourceDataFromDB.iloc[0]["ABN"]
        StoreAddress = SourceDataFromDB.iloc[0]["Address"]
        DefaultBackupPath = "C:\ZiiBackup"
        JobListFormatForPrinter1 = SourceDataFromDB.iloc[0]["JobListFormatForPrinter1"] 
        JobListFormatForPrinter2 = SourceDataFromDB.iloc[0]["JobListFormatForPrinter2"] 
        JobListFormatForPrinter3 = SourceDataFromDB.iloc[0]["JobListFormatForPrinter3"] 
        JobListFormatForPrinter4 = SourceDataFromDB.iloc[0]["JobListFormatForPrinter4"] 
        JobListFormatForPrinter5 = SourceDataFromDB.iloc[0]["JobListFormatForPrinter5"] 
        JobListFormatForPrinter6 = SourceDataFromDB.iloc[0]["JobListFormatForPrinter6"] 
        JobListFormatForPrinter7 = SourceDataFromDB.iloc[0]["JobListFormatForPrinter7"] 
        JobListFormatForPrinter8 = SourceDataFromDB.iloc[0]["JobListFormatForPrinter8"] 
        JobListFormatForPrinter9 = SourceDataFromDB.iloc[0]["JobListFormatForPrinter9"] 
        JobListFormatForPrinter10 = SourceDataFromDB.iloc[0]["JobListFormatForPrinter10"] 
        JobListFormatForPrinter11 = SourceDataFromDB.iloc[0]["JobListFormatForPrinter11"] 
        JobListFormatForPrinter12 = SourceDataFromDB.iloc[0]["JobListFormatForPrinter12"] 


        sqlString = "UPDATE Profile SET CompanyName = N'" + StoreCompanyName + "', ABN = N'"+ StoreABN +" ', Address = N'" + StoreAddress + "', DefaultBackupPath = N'" + DefaultBackupPath +"', JobListFormatForPrinter1 = " + JobListFormatForPrinter1  +", JobListFormatForPrinter2 = " + JobListFormatForPrinter2  +", JobListFormatForPrinter3 = " + JobListFormatForPrinter3  +", JobListFormatForPrinter4 = " + JobListFormatForPrinter4  +", JobListFormatForPrinter5 = " + JobListFormatForPrinter5  +", JobListFormatForPrinter6 = " + JobListFormatForPrinter6  +", JobListFormatForPrinter7 = " + JobListFormatForPrinter7  +", JobListFormatForPrinter8 = " + JobListFormatForPrinter8  +", JobListFormatForPrinter9 = " + JobListFormatForPrinter9  +", JobListFormatForPrinter10 = " + JobListFormatForPrinter10  +", JobListFormatForPrinter11 = " + JobListFormatForPrinter11  +", JobListFormatForPrinter12 = " + JobListFormatForPrinter12 + ";"

        TargetSQLServerCursor.execute(sqlString)
        TargetSQLServerConnection.commit()
        print("Rows Affacted = ",TargetSQLServerCursor.rowcount)

    


        



        # -------------------- table layout ------------------------------------
        print("\n\n table layout")
        # clear current table layout
        truncatetabeSQLTablePage="truncate table TablePage;"
        truncatetabeSQLTableSet="truncate table TableSet;"
        TargetSQLServerCursor.execute(truncatetabeSQLTablePage)
        TargetSQLServerConnection.commit()
        print("Rows Affacted = ",TargetSQLServerCursor.rowcount)
        TargetSQLServerCursor.execute(truncatetabeSQLTableSet)
        TargetSQLServerConnection.commit()
        print("Rows Affacted = ",TargetSQLServerCursor.rowcount)

        #print(truncatetabeSQLTablePage)
        #Rebuild TablePage table
        TablePageListQuery = "SELECT PageNo, Description FROM TablePage"
        TablePageSqlResult = pd.read_sql(TablePageListQuery, SourceSQLServerConnection)
        TablePageList = TablePageSqlResult.astype("string")

        #INSERT INTO [dbo].[TablePage] ([PageNo], [Description]) VALUES (1, N'DINE IN');

        indexOfTablePageList=0
        lengthOfTablePageList = len(TablePageList)

        for indexOfTablePageList in range(lengthOfTablePageList):
            PageNo = TablePageList.iloc[indexOfTablePageList]["PageNo"]
            pageDescription  = TablePageList.iloc[indexOfTablePageList]["Description"]

            TablePageInsertQuery = "INSERT INTO [dbo].[TablePage] ([PageNo], [Description]) VALUES ("+PageNo+", N'"+pageDescription+"');"
            TargetSQLServerCursor.execute(TablePageInsertQuery)
            TargetSQLServerConnection.commit()
            print("Rows Affacted = ",TargetSQLServerCursor.rowcount)

       
        
        #Rebuild TableSet table
        TableSetListQuery = "SELECT * FROM TableSet;"

        TableSetSqlResult = pd.read_sql(TableSetListQuery, SourceSQLServerConnection)
        TableSetList = TableSetSqlResult.astype("string")

        #INSERT INTO [dbo].[TableSet] ([Status], [TableNo], [Seats], [FontName], [FontSize], [FontBold], [FontItalic], [FontUnderline], [FontStrikeout], [ButtonShape], [ButtonWidth], [ButtonHeight], [ButtonX], [ButtonY], [PropertyFlag], [Description], [PageFlag], [PDAPosition], [MinimumChargePerTable], [ServiceStatus], [IPAddress], [SelfOrderStatus], [TerminalConnected], [TableLockerName], [OnlineOrderTable], [PId], [ZiiTOTableLockName], [TeamNo], [LockUpdateTime], [TeamLocalTime]) VALUES (0, N'TA7', 0, N'Tahoma', 18, '1', '0', '0', '0', 3, 60, 50, 1, 334, '1', N'TA7', 2, 48, 0, 0, NULL, '0', '0', NULL, '1', 355, NULL, NULL, NULL, NULL);


        indexOfTableSetList=0
        lengthOfTableSetList = len(TableSetList)

        for indexOfTableSetList in range(lengthOfTableSetList):
            TableNo = TableSetList.iloc[indexOfTableSetList]["TableNo"]
            Description = TableSetList.iloc[indexOfTableSetList]["Description"]
            Seats = TableSetList.iloc[indexOfTableSetList]["Seats"]
            Status = TableSetList.iloc[indexOfTableSetList]["Status"]
            ButtonX = TableSetList.iloc[indexOfTableSetList]["ButtonX"]
            ButtonY = TableSetList.iloc[indexOfTableSetList]["ButtonY"]
            ButtonWidth = TableSetList.iloc[indexOfTableSetList]["ButtonWidth"]
            ButtonHeight = TableSetList.iloc[indexOfTableSetList]["ButtonHeight"]
            ButtonShape = TableSetList.iloc[indexOfTableSetList]["ButtonShape"]
            PropertyFlag = TableSetList.iloc[indexOfTableSetList]["PropertyFlag"]
            PageFlag = TableSetList.iloc[indexOfTableSetList]["PageFlag"]
            PDAPosition = TableSetList.iloc[indexOfTableSetList]["PDAPosition"]
           
            OnlineOrderTable = TableSetList.iloc[indexOfTableSetList]["OnlineOrderTable"]

            if TableSetList.iloc[indexOfTableSetList]["Status"] == '1':
                Description = TableSetList.iloc[indexOfTableSetList]["TableNo"]

         
          
            #print(TableSetList.iloc[indexOfTableSetList])
            
            TableSetinsertQuery = "INSERT INTO [dbo].[TableSet] ([Status], [TableNo], [Seats], [FontName], [FontSize], [FontBold], [FontItalic], [FontUnderline], [FontStrikeout], [ButtonShape], [ButtonWidth], [ButtonHeight], [ButtonX], [ButtonY], [PropertyFlag], [Description], [PageFlag], [PDAPosition], [MinimumChargePerTable], [ServiceStatus], [IPAddress], [SelfOrderStatus], [TerminalConnected], [TableLockerName], [OnlineOrderTable],  [ZiiTOTableLockName], [TeamNo], [LockUpdateTime], [TeamLocalTime]) VALUES ("+Status+", N'"+TableNo+"', "+Seats+", N'Tahoma', 18, '1', '0', '0', '0', "+ButtonShape+", "+ButtonWidth+", "+ButtonHeight+", "+ButtonX+", "+ButtonY+", '"+PropertyFlag+"', N'"+ Description +"', "+PageFlag+", "+PDAPosition+", 0, 0, NULL, '0', '0', NULL, '"+OnlineOrderTable+"', NULL, NULL, NULL, NULL);"
            print(Description)
            #print(TableNo)
            #print(TableSetinsertQuery)
            TargetSQLServerCursor.execute(TableSetinsertQuery)
            TargetSQLServerConnection.commit()
            #print("Rows Affacted = ",TargetSQLServerCursor.rowcount)


        
        
        

        # ----------------------------------------- User Account ---------------------------------------------------------------------
        print("\n\n User Account")
        StaffListQuery = "SELECT StaffName, SecureCode FROM AccessMenu"

        SqlResult2 = pd.read_sql(StaffListQuery, SourceSQLServerConnection)
        StaffList = SqlResult2.astype("string")

        x=0
        lengthOfStaffList = len(StaffList)



        for x in range(lengthOfStaffList):
            #print(StaffList[x]["StaffName"])

            if  StaffList.iloc[x]["StaffName"]=='Z OKAMI':
                print(StaffList.iloc[x]["StaffName"] )

            elif StaffList.iloc[x]["StaffName"]=='STAFF':
                print(StaffList.iloc[x]["StaffName"] )

            elif StaffList.iloc[x]["StaffName"]=='ZIITECH':
                print(StaffList.iloc[x]["StaffName"] )

            elif StaffList.iloc[x]["StaffName"]=='SUPERVISOR':
                print(StaffList.iloc[x]["StaffName"] )
            
            elif StaffList.iloc[x]["StaffName"]=='OKAMI':
                print(StaffList.iloc[x]["StaffName"] )

            elif StaffList.iloc[x]["StaffName"]=='lotus':
                print(StaffList.iloc[x]["StaffName"] )

            else :
                staffName = StaffList.iloc[x]["StaffName"]
                staffPassword  = StaffList.iloc[x]["SecureCode"]

                insertQuery = "INSERT INTO [dbo].[AccessMenu] ([AuthoriseCloseWindow], [AuthoriseDiscount], [BookingFormConditionSetupMenu], [DailyReportMenu], [DatabaseBackupMenu], [DatabaseRestoreMenu], [InvoiceConditionSetupMenu], [OpenCashDrawerMenu], [PaymentAuthority], [PrintInvoiceAuthority], [PrintJobListAuthority], [TableInformationSetupMenu], [StaffName], [SecureCode], [Supervisor], [BookingListMenu], [StockReceiveMenu], [InquirySalesHistoryMenu], [VIPInformationMenu], [SalesReportMenu], [SalesStatisticsReportMenu], [StockReportMenu], [StockReceiveReportMenu], [StatisticsChartMenu], [SupplierInformationListMenu], [ExpensesDescriptionSetupMenu], [ExpensesDataEntryMenu], [ExpensesReportMenu], [ReceiptsReportMenu], [PaymentsReportMenu], [GSTPayableReportMenu], [ProfileSetupMenu], [PrinterSetupMenu], [CategorySetupMenu], [MenuSetupMenu], [PaymentsMethodSetupMenu], [SupplierInformationSetupMenu], [Birthday], [Telephone], [Mobile], [Fax], [Address], [Rate], [AttendanceReportMenu], [VoidItemAuthority], [PurchaseOrderMenu], [PurchasePayableMenu], [TableOrderMenu], [PointOfSalesMenu], [CheckDailyReport], [AuthoriseRefund], [UserManager], [AllowEditOrder], [PrintDailyReport], [DrawerPortNumber], [DefaultDrawerPortNumber], [EditAttendanceRecord], [StockAdjustmentMenu], [StockAdjustmentReportMenu], [PhoneOrderMenu], [CashPayOutMenu], [CashFloatMenu], [AssignDriverAuthorised], [DepositMenu], [WastageMenu], [WastageReportMenu], [AuthrisedCancelHoldOrder], [ManuallyEnterDiscountRate], [EditOrderPayment], [InquirySalesRelatedReportDays], [CashDeclarationReportMenu], [AccountEnabled], [AuthorizedChangeQty], [AuthorizedChangePrice], [DeleteVIPRecord], [ControlButtonSetup], [DiscountRateSetup], [VoidItemDescriptionSetup], [EFTPOSUtility], [ChangeMenuStatus], [StockTakeMenu], [StockTakeReportMenu], [UserGroupSetupAuthorized], [UploadMembersRewardsMenu],  [SettingsPortalMenu], [ZiiTOTableLockMenu], [StaffCode], [LastUpdatedTime], [FirstName], [LastName]) VALUES ('0', '0', '0', '0', '0', '0', '0', '0', '0', '1', '1', '0', N'"+staffName+"', N'"+ staffPassword  +"', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '2019-07-01 00:00:00.000', N'', N'', N'', N'', 0, '0', '0', '0', '0', '1', '1', '0', '0', '0', '0', '0', NULL, NULL, '0', '0', '0', '1', '0', '0', '0', '0', '0', '0', '0', '0', '0', 0, '0', '1', '0', '0', '0', '0', '0', '0', '1', '0', '0', '0', '0', '0', '0', '1', NEWID(), '2022-02-14 22:06:28.277', N'"+staffName+"', N'');"

                TargetSQLServerCursor.execute(insertQuery)
                TargetSQLServerConnection.commit()
                #print("Staff Rows Affacted = ",TargetSQLServerCursor.rowcount)


      
        SourceSQLServerConnection.close()
        TargetSQLServerConnection.close()
        print("\n")
        print("Go to Next")



    print(".............All Process completed.............")



main()

