import sys
import os
import pyodbc 
import pandas as pd
import numpy as np
import signal




def main():
    
    
    #   -------------------  Configurations -------------------------------
    #                     ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
    server = '192.168.1.85,9899'
    database = '41OKAMI_Newton'
    username = 'sa'
    password = '0000'

    #Password Connection
    PassSQLServerConnection = pyodbc.connect('DRIVER={SQL Server}; SERVER='+server+'; DATABASE='+database+'; UID='+username+'; PWD='+ password)
    #Windows Auth
    #WindowsAuthSQLServerConnection = pyodbc.connect('DRIVER={SQL Server}; SERVER='+server+'; DATABASE='+database+'; Trusted_Connection=True;' )

    #cursor = PassSQLServerConnection.cursor()
    
    MenuItemQuery = "SELECT  ItemCode ,Description1 ,Description2 ,Category ,PrinterPort ,PrinterPort1 ,PrinterPort2 ,PrinterPort3 FROM MenuItem order by ItemCode"
    
    
    OutputFileName = database + '_outputFile.xlsx';
    
    #                        ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑            
    #   -------------------  Configurations -------------------------------
    
    
    
    
    
    
    
  

    SqlResult1 = pd.read_sql(MenuItemQuery, PassSQLServerConnection)
    SourceDataFromDB = SqlResult1.astype("string")
    NewExcelDataFrame = pd.DataFrame()
  

    #print("total MenuItem rows: ")
    #print(len(SourceDataFromDB))

    #print("total rows: ")
    #print(len(NewExcelDataFrame))

    #print("Export to new excel")
    #SourceDataFromDB.to_excel(r'export_dataframe.xlsx', index = True, header=True)
    #print("Process completed")
    
    
    print("Start convert to ZiiPOS")
    
    ZiiPOSExcelTemplete = pd.read_excel('OKAMI_STANDER_V1.xlsx', index_col=None,dtype = str)
    ZiiPOSExcel=ZiiPOSExcelTemplete.astype("string")
    ZiiPOSExcel_Done = processPrinterSetting(ZiiPOSExcel,SourceDataFromDB)
    
    
    #print(ZiiPOSExcel_Done)
    ZiiPOSExcel_Done.to_excel(OutputFileName, index = True, header=True)





def processPrinterSetting(ZiiPOSExcel, SourceDataFromDB):
    
    x=0
    
    
    NewExcelDataFrame=pd.DataFrame()
    
    
        
    for x in range(len(ZiiPOSExcel)):
        #print(x ," / ",len(ZiiPOSExcel))
        prograss=round( x/len(ZiiPOSExcel)*100,1)
        
        if (prograss>30.0 and prograss<30.2):
            print(str(prograss)+'%')
        
        elif (prograss>60 and prograss<60.2):
            print(str(prograss)+'%')
            
        elif (prograss>90 and prograss<90.2):
            print(str(prograss)+'%')
        
        tempItemCode = ZiiPOSExcel.iloc[x]["ItemCode"]     
        tempReadData = ZiiPOSExcel.iloc[x]
        #ItemCode ,Description1 ,Description2 ,Category ,PrinterPort ,PrinterPort1 ,PrinterPort2 ,PrinterPort3
            
        # tempReadData["ItemCode"]=ZiiPOSExcel.iloc[x]["ItemCode"]
        # tempReadData["Description1"]=ZiiPOSExcel.iloc[x]["Description1"]
        # tempReadData["Description2"]=ZiiPOSExcel.iloc[x]["Description2"]
        # NewExcelDataFrame = NewExcelDataFrame.append (tempReadData)
        # tempReadData["PrinterPort"]=ZiiPOSExcel.iloc[x]["PrinterPort"]
        # tempReadData["PrinterPort1"]=ZiiPOSExcel.iloc[x]["PrinterPort1"]
        # tempReadData["PrinterPort2"]=ZiiPOSExcel.iloc[x]["PrinterPort2"]
        # tempReadData["PrinterPort3"]=ZiiPOSExcel.iloc[x]["PrinterPort3"]
        # tempReadData["Category"]=ZiiPOSExcel.iloc[x]["Category"]     
       
       
        tempResult=SourceDataFromDB.loc[SourceDataFromDB['ItemCode']== tempItemCode]
        
            
        if ZiiPOSExcel.iloc[x]["ItemCode"] =='BA01':
            
           tempResult_SIGNATURE_SET=SourceDataFromDB.loc[SourceDataFromDB['ItemCode']== "BA04"] 
           #tempResult_SIGNATURE_SET=SourceDataFromDB.loc[SourceDataFromDB['Description1'].str.contains("SIGNATURE", case=False)]
           tempReadData["PrinterPort"]=tempResult_SIGNATURE_SET["PrinterPort"].item()
           tempReadData["PrinterPort1"]=tempResult_SIGNATURE_SET["PrinterPort1"].item()
           tempReadData["PrinterPort2"]=tempResult_SIGNATURE_SET["PrinterPort2"].item()
           tempReadData["PrinterPort3"]=tempResult_SIGNATURE_SET["PrinterPort3"].item()
           
        elif ZiiPOSExcel.iloc[x]["ItemCode"] =='BA06':
            #FATHER'S DAY SET

            tempResult_FDS=SourceDataFromDB.loc[SourceDataFromDB['ItemCode']== "BA04"]
            tempReadData["PrinterPort"]=tempResult_FDS["PrinterPort"].item()
            tempReadData["PrinterPort1"]=tempResult_FDS["PrinterPort1"].item()
            tempReadData["PrinterPort2"]=tempResult_FDS["PrinterPort2"].item()
            tempReadData["PrinterPort3"]=tempResult_FDS["PrinterPort3"].item()
       
        elif ZiiPOSExcel.iloc[x]["ItemCode"] =='BB00':
            tempReadData["PrinterPort"]="0"
            tempReadData["PrinterPort1"]="0"
            tempReadData["PrinterPort2"]="0"
            tempReadData["PrinterPort3"]="0"
           
        elif ZiiPOSExcel.iloc[x]["ItemCode"] =='121':
            #Instrction how many pieces totally

            tempReadData["PrinterPort"]="0"
            tempReadData["PrinterPort1"]="0"
            tempReadData["PrinterPort2"]="0"
            tempReadData["PrinterPort3"]="0"
            
            
        elif ZiiPOSExcel.iloc[x]["ItemCode"] =='L530':
            #L530	TAKEAWAY BAG

            tempReadData["PrinterPort"]="0"
            tempReadData["PrinterPort1"]="0"
            tempReadData["PrinterPort2"]="0"
            tempReadData["PrinterPort3"]="0"
            
            
            
  
         
            
            
            
        
            
            
        elif ZiiPOSExcel.iloc[x]["ItemCode"] =='C403':
            #C403	(TA) SHSHI ROLL PLATTER
            tempResult_C401=SourceDataFromDB.loc[SourceDataFromDB['ItemCode']== "C401"]
            tempReadData["PrinterPort"]=tempResult_C401["PrinterPort"].item()
            tempReadData["PrinterPort1"]=tempResult_C401["PrinterPort1"].item()
            tempReadData["PrinterPort2"]=tempResult_C401["PrinterPort2"].item()
            tempReadData["PrinterPort3"]=tempResult_C401["PrinterPort3"].item()
       
            
        #  Instractions     
        elif ZiiPOSExcel.iloc[x]["ItemCode"] =='$001':
            tempReadData["PrinterPort"]="0"
            tempReadData["PrinterPort1"]="0"
            tempReadData["PrinterPort2"]="0"
            tempReadData["PrinterPort3"]="0"
            
        #  Instractions   
        elif ZiiPOSExcel.iloc[x]["ItemCode"] =='_001':
            tempReadData["PrinterPort"]="0"
            tempReadData["PrinterPort1"]="0"
            tempReadData["PrinterPort2"]="0"
            tempReadData["PrinterPort3"]="0"
            
        #  Instractions   
            
        elif ZiiPOSExcel.iloc[x]["ItemCode"] =='_002':
            tempReadData["PrinterPort"]="0"
            tempReadData["PrinterPort1"]="0"
            tempReadData["PrinterPort2"]="0"
            tempReadData["PrinterPort3"]="0"
            
        #  Instractions   
            
        elif ZiiPOSExcel.iloc[x]["ItemCode"] =='_003':
            tempReadData["PrinterPort"]="0"
            tempReadData["PrinterPort1"]="0"
            tempReadData["PrinterPort2"]="0"
            tempReadData["PrinterPort3"]="0"
            
            
        #  Instractions   
            
        elif ZiiPOSExcel.iloc[x]["ItemCode"] =='_004':
            tempReadData["PrinterPort"]="0"
            tempReadData["PrinterPort1"]="0"
            tempReadData["PrinterPort2"]="0"
            tempReadData["PrinterPort3"]="0"
            
         #  Instractions   
            
        elif ZiiPOSExcel.iloc[x]["ItemCode"] =='_005':
            tempReadData["PrinterPort"]="0"
            tempReadData["PrinterPort1"]="0"
            tempReadData["PrinterPort2"]="0"
            tempReadData["PrinterPort3"]="0"
            
            
        elif ZiiPOSExcel.iloc[x]["ItemCode"] =='TI13':
                
             #FSUSHI NIGIRI PLATTER( SET ITEM)


            tempResult_TI13=SourceDataFromDB.loc[SourceDataFromDB['Description1']== "VEG GYOZA (SET ITEM)"]
            tempReadData["PrinterPort"]=tempResult_TI13["PrinterPort"].item()
            tempReadData["PrinterPort1"]=tempResult_TI13["PrinterPort1"].item()
            tempReadData["PrinterPort2"]=tempResult_TI13["PrinterPort2"].item()
            tempReadData["PrinterPort3"]=tempResult_TI13["PrinterPort3"].item()
            
        elif ZiiPOSExcel.iloc[x]["ItemCode"] =='TI31':
            
             #FSUSHI NIGIRI PLATTER( SET ITEM)


            tempResult_TI20=SourceDataFromDB.loc[SourceDataFromDB['ItemCode']== "TI20"]
            tempReadData["PrinterPort"]=tempResult_TI20["PrinterPort"].item()
            tempReadData["PrinterPort1"]=tempResult_TI20["PrinterPort1"].item()
            tempReadData["PrinterPort2"]=tempResult_TI20["PrinterPort2"].item()
            tempReadData["PrinterPort3"]=tempResult_TI20["PrinterPort3"].item()
       
       
        elif ZiiPOSExcel.iloc[x]["ItemCode"] =='TI32':
             #MINI MATCHA TAIYAKI(SET ITEM)

            tempResultA200=SourceDataFromDB.loc[SourceDataFromDB['ItemCode']== "A200"]
            tempReadData["PrinterPort"]=tempResultA200["PrinterPort"].item()
            tempReadData["PrinterPort1"]=tempResultA200["PrinterPort1"].item()
            tempReadData["PrinterPort2"]=tempResultA200["PrinterPort2"].item()
            tempReadData["PrinterPort3"]=tempResultA200["PrinterPort3"].item()
            
            
        elif ZiiPOSExcel.iloc[x]["ItemCode"] =='A229':
            #208.Kaki Fry
            tempResultA211=SourceDataFromDB.loc[SourceDataFromDB['ItemCode']== "A211"]
            tempReadData["PrinterPort"]=tempResultA211["PrinterPort"].item()
            tempReadData["PrinterPort1"]=tempResultA211["PrinterPort1"].item()
            tempReadData["PrinterPort2"]=tempResultA211["PrinterPort2"].item()
            tempReadData["PrinterPort3"]=tempResultA211["PrinterPort3"].item()
                    
     

        elif tempResult.empty:
                
            if ZiiPOSExcel.iloc[x]["Category"] =='[A] RICE +NODDLES':
                tempResultJ451=SourceDataFromDB.loc[SourceDataFromDB['ItemCode']== "J451"]
                tempReadData["PrinterPort"]=tempResultJ451["PrinterPort"].item()
                tempReadData["PrinterPort1"]=tempResultJ451["PrinterPort1"].item()
                tempReadData["PrinterPort2"]=tempResultJ451["PrinterPort2"].item()
                tempReadData["PrinterPort3"]=tempResultJ451["PrinterPort3"].item()
                
            elif ZiiPOSExcel.iloc[x]["Category"] =='[A] BENTO + DESSERT':
                tempResultK501=SourceDataFromDB.loc[SourceDataFromDB['ItemCode']== "K501"]
                tempReadData["PrinterPort"]=tempResultK501["PrinterPort"].item()
                tempReadData["PrinterPort1"]=tempResultK501["PrinterPort1"].item()
                tempReadData["PrinterPort2"]=tempResultK501["PrinterPort2"].item()
                tempReadData["PrinterPort3"]=tempResultK501["PrinterPort3"].item()
                
                
            elif ZiiPOSExcel.iloc[x]["Category"] =='INSTU':
                
                tempResultQ001=SourceDataFromDB.loc[SourceDataFromDB['ItemCode']== "Q001"]
                tempReadData["PrinterPort"]=tempResultQ001["PrinterPort"].item()
                tempReadData["PrinterPort1"]=tempResultQ001["PrinterPort1"].item()
                tempReadData["PrinterPort2"]=tempResultQ001["PrinterPort2"].item()
                tempReadData["PrinterPort3"]=tempResultQ001["PrinterPort3"].item()
                
            elif ZiiPOSExcel.iloc[x]["Category"] =='SOFT DRINK+TEA+COFFEE':
                tempResult555A=SourceDataFromDB.loc[SourceDataFromDB['ItemCode']== "555A"]
                tempReadData["PrinterPort"]=tempResult555A["PrinterPort"].item()
                tempReadData["PrinterPort1"]=tempResult555A["PrinterPort1"].item()
                tempReadData["PrinterPort2"]=tempResult555A["PrinterPort2"].item()
                tempReadData["PrinterPort3"]=tempResult555A["PrinterPort3"].item()

            elif ZiiPOSExcel.iloc[x]["Category"] =='[A] SUSHI&SASHIMI':
                tempResultB251=SourceDataFromDB.loc[SourceDataFromDB['ItemCode']== "B251"]
                tempReadData["PrinterPort"]=tempResultB251["PrinterPort"].item()
                tempReadData["PrinterPort1"]=tempResultB251["PrinterPort1"].item()
                tempReadData["PrinterPort2"]=tempResultB251["PrinterPort2"].item()
                tempReadData["PrinterPort3"]=tempResultB251["PrinterPort3"].item()
                
            elif ZiiPOSExcel.iloc[x]["Category"] =='622~627 SPIRIT':
                tempResultPI01=SourceDataFromDB.loc[SourceDataFromDB['ItemCode']== "PI01"]
                tempReadData["PrinterPort"]=tempResultPI01["PrinterPort"].item()
                tempReadData["PrinterPort1"]=tempResultPI01["PrinterPort1"].item()
                tempReadData["PrinterPort2"]=tempResultPI01["PrinterPort2"].item()
                tempReadData["PrinterPort3"]=tempResultPI01["PrinterPort3"].item()
                
            elif ZiiPOSExcel.iloc[x]["Category"] =='BEER + SAKE +PLUM WINE':
                tempResultN563=SourceDataFromDB.loc[SourceDataFromDB['ItemCode']== "N563"]
                tempReadData["PrinterPort"]=tempResultN563["PrinterPort"].item()
                tempReadData["PrinterPort1"]=tempResultN563["PrinterPort1"].item()
                tempReadData["PrinterPort2"]=tempResultN563["PrinterPort2"].item()
                tempReadData["PrinterPort3"]=tempResultN563["PrinterPort3"].item()
                
            elif ZiiPOSExcel.iloc[x]["Category"] =='SOFT DRINK+TEA+COFFEE':
                tempResult555A=SourceDataFromDB.loc[SourceDataFromDB['ItemCode']== "555A"]
                tempReadData["PrinterPort"]=tempResult555A["PrinterPort"].item()
                tempReadData["PrinterPort1"]=tempResult555A["PrinterPort1"].item()
                tempReadData["PrinterPort2"]=tempResult555A["PrinterPort2"].item()
                tempReadData["PrinterPort3"]=tempResult555A["PrinterPort3"].item()
                
            elif ZiiPOSExcel.iloc[x]["Category"] =='[B] DESSERT':
                tempResult191=SourceDataFromDB.loc[SourceDataFromDB['ItemCode']== "191"]
                tempReadData["PrinterPort"]=tempResult191["PrinterPort"].item()
                tempReadData["PrinterPort1"]=tempResult191["PrinterPort1"].item()
                tempReadData["PrinterPort2"]=tempResult191["PrinterPort2"].item()
                tempReadData["PrinterPort3"]=tempResult191["PrinterPort3"].item()    
         
            else:
         
                tempReadData["PrinterPort"]="9999"
                tempReadData["PrinterPort1"]="0"
                tempReadData["PrinterPort2"]="0"
                tempReadData["PrinterPort3"]="0"
                
             
                 
      
            
        else:
            
            tempReadData["PrinterPort"]=tempResult["PrinterPort"].item()
            tempReadData["PrinterPort1"]=tempResult["PrinterPort1"].item()
            tempReadData["PrinterPort2"]=tempResult["PrinterPort2"].item()
            tempReadData["PrinterPort3"]=tempResult["PrinterPort3"].item()
                
           
           
        NewExcelDataFrame = NewExcelDataFrame.append (tempReadData)
            
            
           
            

    
        
    return NewExcelDataFrame











main()

