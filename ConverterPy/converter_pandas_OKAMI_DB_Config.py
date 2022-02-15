import sys
import os
import pyodbc 
import pandas as pd





def main():
    
    
    #   -------------------  Configurations -------------------------------
    #                        ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
    server = 'localhost,9899'
    sourceDatabase = '19Okami_Chelsea_Heights'
  
    username = 'ZiiPos'
    password = 'ZiiPos884568'

    #Password Connection
    SourceSQLServerConnection = pyodbc.connect('DRIVER={SQL Server}; SERVER='+server+'; DATABASE='+sourceDatabase+'; UID='+username+'; PWD='+ password)

    
    excelOutput=pd.DataFrame()
    xcelOutput=pd.DataFrame()
    outputFile = open('sqlstringOutput.txt', 'a')



    infoQuery = "Select CompanyName, Telephone,ABN,Address,JobListFormatForPrinter1,JobListFormatForPrinter2,JobListFormatForPrinter3,JobListFormatForPrinter4,JobListFormatForPrinter5,JobListFormatForPrinter6,JobListFormatForPrinter7,JobListFormatForPrinter8,JobListFormatForPrinter9,JobListFormatForPrinter10,JobListFormatForPrinter11,JobListFormatForPrinter12 From Profile;"
    
    SqlResult1 = pd.read_sql(infoQuery, SourceSQLServerConnection)
    SourceDataFromDB = SqlResult1.astype("string")



    StaffListQuery = "SELECT StaffName, SecureCode FROM AccessMenu"

    SqlResult2 = pd.read_sql(StaffListQuery, SourceSQLServerConnection)
    StaffList = SqlResult2.astype("string")





    
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


    sqlString = "UPDATE Profile SET CompanyName = N'" + StoreCompanyName + "', Telephone = N'" + StoreTelephone+"', ABN = N'"+ StoreABN +" ', Address = N'" + StoreAddress + "', DefaultBackupPath = N'" + DefaultBackupPath +"', JobListFormatForPrinter1 = " + JobListFormatForPrinter1  +", JobListFormatForPrinter2 = " + JobListFormatForPrinter2  +", JobListFormatForPrinter3 = " + JobListFormatForPrinter3  +", JobListFormatForPrinter4 = " + JobListFormatForPrinter4  +", JobListFormatForPrinter5 = " + JobListFormatForPrinter5  +", JobListFormatForPrinter6 = " + JobListFormatForPrinter6  +", JobListFormatForPrinter7 = " + JobListFormatForPrinter7  +", JobListFormatForPrinter8 = " + JobListFormatForPrinter8  +", JobListFormatForPrinter9 = " + JobListFormatForPrinter9  +", JobListFormatForPrinter10 = " + JobListFormatForPrinter10  +", JobListFormatForPrinter11 = " + JobListFormatForPrinter11  +", JobListFormatForPrinter12 = " + JobListFormatForPrinter12 + ";"


    outputFile.write(sqlString)

    #outputFile.write(str(sqlString))
    outputFile.write('\n')
    outputFile.write('\n')


    


    #print(len(StaffList))

    x=0
    lengthOfStaffList = len(StaffList)

    #print(StaffList[1]["StaffName"])

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

            #print(insertQuery)
            outputFile.write('\n')
            outputFile.write(insertQuery)
            outputFile.write('\n')
           

    #print(StaffList)
    

       
       
    

    
    




main()

