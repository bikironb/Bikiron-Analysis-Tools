# -*- coding: utf-8 -*-
"""
Created on Thu Sep 22 13:13:40 2022

@author: BikironBanerjee
"""

import xlwings as xw
import pandas as pd 
import numpy as np
from pathlib import Path
import xlwings as xw
import os as o
from matplotlib import pyplot as plt

# parameritise excel to create each template
# for each template get start and end variables for each set
# code in alignment data sections


# make general info dependent on name of header
# error handling


def main(params):
    wb = xw.Book.caller()
    
    ws_FactsheetTemplate = wb.sheets['Factsheet Template']
    ws_Data  = wb.sheets['Activities Data']
    ws_Piechartdata = wb.sheets['PieChartDataSource']
        
    last_row_MetricData = wb.sheets[ws_Data].range(7,1).end('down').row
        
    data = ws_Data.range((7,1),(last_row_MetricData,150)).value
    dfData = pd.DataFrame(data)

    dfData = dfData.rename(columns=dfData.iloc[0]).drop(dfData.index[0])

    # get start and end rows for each section
    startUOP = 0
    endUOP = 0
    startPCT = 0
    endPCT = 0
    startActName = 0
    endActName = 0

    #get column range for activities
    for (columnName, columnData) in dfData.iteritems():
         if columnName[0:3] == "UOP": # look at prefix of col name - UOP, use of proceeds
             if startUOP == 0:
                 startUOP = dfData.columns.get_loc(columnName)
         
         if columnName[0:3] == "PCT": # look at prefix of col name - PCT - Percentage exposure
             if endUOP == 0:
                 endUOP = dfData.columns.get_loc(columnName)
                 startPCT = endUOP + 1

         if columnName[0:3] == "Act": # look at prefix of col name - PCT - Percentage exposure
             if endPCT == 0:
                 endPCT = dfData.columns.get_loc(columnName)
                 startActName = endPCT
         
         if columnName[0:3] == "END": # look at prefix of col name - PCT - Percentage exposure
             if endActName == 0:
                endActName = dfData.columns.get_loc(columnName)

    #loop though data , write to sheet and create pie charts             


    # make completely user driven - take field names from config table
    # - make strFieldname1, 2 etc for each general data item - loop through general data section - make a list of all items
    # place all items in list by looping list and running place statements

    
    for row in dfData.index:
        
        datarow = row-1
        
        ISINloc = dfData.columns.get_loc('ISIN')
        strISIN = dfData.iloc[datarow,ISINloc]          #ISIN
        
        IssuerNameLoc = dfData.columns.get_loc('Name')
        strIssuerName = dfData.iloc[datarow,IssuerNameLoc]    #Issuer Name
        
        LabelLoc = dfData.columns.get_loc('Labelling')
        strLabel = dfData.iloc[datarow,LabelLoc]         #Label
        
        MatDateLoc = dfData.columns.get_loc('Final Maturity')
        strMaturityDate = dfData.iloc[datarow,MatDateLoc]  #Maturity Date
    
        EligLoc = dfData.columns.get_loc('Eligibility')
        fltEligibility = dfData.iloc[datarow,EligLoc]   #Eligibility
    
        ESGRatingLoc = dfData.columns.get_loc('ESG Rating')
        fltESGRating = dfData.iloc[datarow,ESGRatingLoc]     #ESG Rating
    
        AmountLoc = dfData.columns.get_loc('Amount issued')
        fltAmount = dfData.iloc[datarow,AmountLoc]     #ESG Rating
    
        # Place ISSUER
        strCellIssuer = ws_FactsheetTemplate.range("W27").value
        ws_FactsheetTemplate.range(strCellIssuer).options(index=False, header=True).value = strIssuerName
    
        # Place ISIN
        strCellISIN = ws_FactsheetTemplate.range("W28").value
        ws_FactsheetTemplate.range(strCellISIN).options(index=False, header=True).value = strISIN        
    
        # Place Labelling
        strCellLabelling = ws_FactsheetTemplate.range("W31").value        
        ws_FactsheetTemplate.range(strCellLabelling).options(index=False, header=True).value = strLabel        
    
        # Place ESG Rating
        strCellRating = ws_FactsheetTemplate.range("W32").value        
        ws_FactsheetTemplate.range("F32").options(index=False, header=True).value = fltESGRating        
    
        # Place Eligibility
        strCellElig = ws_FactsheetTemplate.range("W33").value        
        ws_FactsheetTemplate.range("F36").options(index=False, header=True).value = fltEligibility
        
        # Place Mat Date
        strCellMatDate = ws_FactsheetTemplate.range("W29").value        
        ws_FactsheetTemplate.range("F20").options(index=False, header=True).value = strMaturityDate
        
        # Amount Issued
        strCellAmount = ws_FactsheetTemplate.range("W30").value        
        ws_FactsheetTemplate.range("F24").options(index=False, header=True).value = fltAmount
    
    
        # select top 4 activity exposures
                    
        dfCurrentISINRowActivities = dfData.iloc[datarow,startUOP:endUOP]
        dfCurrentISINRowActivities = dfCurrentISINRowActivities.to_frame()

        dfCurrentISINRowActivities.sort_values(by=[1],ascending=False,inplace=True)
        
        # sum of all other values
        lastrow = len(dfCurrentISINRowActivities.index)
        SumOther = dfCurrentISINRowActivities.iloc[4:lastrow].sum(axis=0)
        dfSumOther = SumOther.to_frame()

# sum rows 5 - end
#        dfCurrentISINRowActivities[1].apply(pd.to_numeric)
#        dfCurrentISINRowActivities[1] = dfCurrentISINRowActivities[1].replace('',0)
#        dfCurrentISINRowActivities[1] = dfCurrentISINRowActivities[1].replace(np.nan,0)
#        dfCurrentISINRowActivities.drop(dfCurrentISINRowActivities.index[dfCurrentISINRowActivities[1] == 0], inplace=True)
#        dfSortedActivityExposures = dfCurrentISINRowActivities.sort_values(by=[2],ascending=False).head(4)
    
        #make index labels a data column in the dataframe
    #    dfSortedActivityExposures = dfSortedActivityExposures.to_frame()
    
        dfCurrentISINRowActivities.reset_index(inplace=True)
    #    dfSortedActivityExposures = dfSortedActivityExposures[1].apply(pd.to_numeric)
    #    dfSortedActivityExposures = dfSortedActivityExposures.to_frame()
    #    dfSortedActivityExposures = dfSortedActivityExposures*100
    #    dfSortedActivityExposures = dfSortedActivityExposures.drop('index',axis=1)

        ws_Piechartdata.range("Y1").options(index=False, header=False).value = dfCurrentISINRowActivities.head(4)
        ws_Piechartdata.range("B5").options(index=False, header=False).value = dfSumOther
    
        # Alignment Section Code Here
        #loop thorugh PCT and Act - first 2 numbers and first 2 descs
        exp1 = 0
        exp2 = 0
        exp3 = 0
        exp4 = 0
        exp5 = 0
        exp6 = 0
        name1 = ""
        name2 = ""
        name3 = ""
        name4 = ""
        name5 = ""
        name6 = ""
        colName1 = ""
        colName2 = ""
        colName3 = ""
        colName4 = ""
        colName5 = ""
        colName6 = ""
        
#        dfExp = dfData.iloc[0,startPCT:endPCT]
        
#        get exposures
#        dfData[startPCT:endPCT] = dfData[startPCT:endPCT].replace('',0)
#        dfData[startPCT:endPCT] = dfData[startPCT:endPCT].replace(np.nan,0)
#        dfData[startPCT:endPCT].fillna(0, inplace = True)
            
        for i in range(startPCT,endPCT):
            if (dfData.iloc[datarow,i] > 0 ) & (exp1 == 0):
                exp1 = dfData.iloc[datarow,i]
                continue
            if (dfData.iloc[datarow,i] > 0) & (exp2 == 0):
                exp2 = dfData.iloc[datarow,i]
                continue
            if (dfData.iloc[datarow,i] > 0) & (exp3 == 0):
                exp3 = dfData.iloc[datarow,i]
                continue
            if (dfData.iloc[datarow,i] > 0) & (exp4 == 0):
                exp4 = dfData.iloc[datarow,i]    
                continue
            if (dfData.iloc[datarow,i] > 0) & (exp5 == 0):
                exp5 = dfData.iloc[datarow,i]
                continue
            if (dfData.iloc[datarow,i] > 0) & (exp6 == 0):
                exp6 = dfData.iloc[datarow,i]
                continue

#        get Activity Names            
#        dfData[startActName:endActName] = dfData[startActName:endActName].replace('','-')
#        dfData[startActName:endActName] = dfData[startActName:endActName].replace(np.nan,'-')
#        dfData[startActName:endActName].fillna("-", inplace = True)
        
        for i in range(startActName,endActName):
    
            strName = dfData.iloc[datarow,i]    
            if (len(strName) > 2):
                if name1 == "":
                    name1 = dfData.iloc[datarow,i]
                    colName1 = dfData.columns[i]
                    colName1 = colName1.lstrip(colName1[0:8])
                elif (name2 == "") & (name1 != ""):
                    name2 = dfData.iloc[datarow,i]
                    colName2 = dfData.columns[i]
                    colName2 = colName2.lstrip(colName2[0:8])
                elif (name3 == "") & (name2 != ""):
                    name3 = dfData.iloc[datarow,i]
                    colName3 = dfData.columns[i]
                    colName3 = colName3.lstrip(colName3[0:8])
                elif (name4 == "") & (name3 != ""):
                    name4 = dfData.iloc[datarow,i]
                    colName4 = dfData.columns[i]
                    colName4 = colName4.lstrip(colName4[0:8])
                elif (name5 == "") & (name4 != ""):
                    name5 = dfData.iloc[datarow,i]
                    colName5 = dfData.columns[i]
                    colName5 = colName5.lstrip(colName5[0:8])
                elif (name6 == "") & (name5 != ""):
                    name6 = dfData.iloc[datarow,i]
                    colName6 = dfData.columns[i]
                    colName6 = colName6.lstrip(colName6[0:8])


        #place values in sheet - need at least 2 activities
        
        strCellColName1 = ws_FactsheetTemplate.range("W34").value                
        ws_FactsheetTemplate.range(strCellColName1).options(index=False, header=True).value = colName1
        strCellColName2 = ws_FactsheetTemplate.range("W37").value                
        ws_FactsheetTemplate.range(strCellColName2).options(index=False, header=True).value = colName2
        strCellColName3 = ws_FactsheetTemplate.range("W40").value                
        ws_FactsheetTemplate.range(strCellColName3).options(index=False, header=True).value = colName3
        strCellColName4 = ws_FactsheetTemplate.range("W43").value                
        ws_FactsheetTemplate.range(strCellColName4).options(index=False, header=True).value = colName4
        strCellColName5 = ws_FactsheetTemplate.range("W46").value                
        ws_FactsheetTemplate.range(strCellColName5).options(index=False, header=True).value = colName5
        strCellColName6 = ws_FactsheetTemplate.range("W49").value                
        ws_FactsheetTemplate.range(strCellColName6).options(index=False, header=True).value = colName6

        strCellexp1 = ws_FactsheetTemplate.range("W35").value                    
        ws_FactsheetTemplate.range(strCellexp1).options(index=False, header=True).value = exp1
        strCellexp2 = ws_FactsheetTemplate.range("W38").value
        ws_FactsheetTemplate.range(strCellexp2).options(index=False, header=True).value = exp2
        strCellexp3 = ws_FactsheetTemplate.range("W41").value
        ws_FactsheetTemplate.range(strCellexp3).options(index=False, header=True).value = exp3
        strCellexp4 = ws_FactsheetTemplate.range("W44").value
        ws_FactsheetTemplate.range(strCellexp4).options(index=False, header=True).value = exp4
        strCellexp5 = ws_FactsheetTemplate.range("W47").value
        ws_FactsheetTemplate.range(strCellexp5).options(index=False, header=True).value = exp5
        strCellexp6 = ws_FactsheetTemplate.range("W50").value
        ws_FactsheetTemplate.range(strCellexp6).options(index=False, header=True).value = exp6
    
        strCellname1 = ws_FactsheetTemplate.range("W36").value                    
        ws_FactsheetTemplate.range(strCellname1).options(index=False, header=True).value = name1
        strCellname2 = ws_FactsheetTemplate.range("W39").value                    
        ws_FactsheetTemplate.range(strCellname2).options(index=False, header=True).value = name2
        strCellname3 = ws_FactsheetTemplate.range("W42").value                    
        ws_FactsheetTemplate.range(strCellname3).options(index=False, header=True).value = name3
        strCellname4 = ws_FactsheetTemplate.range("W45").value                    
        ws_FactsheetTemplate.range(strCellname4).options(index=False, header=True).value = name4
        strCellname5 = ws_FactsheetTemplate.range("W48").value                    
        ws_FactsheetTemplate.range(strCellname5).options(index=False, header=True).value = name5
        strCellname6 = ws_FactsheetTemplate.range("W51").value                    
        ws_FactsheetTemplate.range(strCellname6).options(index=False, header=True).value = name6


        # get how many further activities to add from config        
        NumberActivities = ws_FactsheetTemplate.range("Y27").value
        
        if NumberActivities == 3:
            strCellColName3 = ws_FactsheetTemplate.range("W40").value                
            ws_FactsheetTemplate.range(strCellColName3).options(index=False, header=True).value = colName3
            
            strCellexp3 = ws_FactsheetTemplate.range("W41").value
            ws_FactsheetTemplate.range(strCellexp3).options(index=False, header=True).value = exp3

            strCellname3 = ws_FactsheetTemplate.range("W42").value                    
            ws_FactsheetTemplate.range(strCellname3).options(index=False, header=True).value = name3

        if NumberActivities == 4:
            strCellColName4 = ws_FactsheetTemplate.range("W43").value                
            ws_FactsheetTemplate.range(strCellColName4).options(index=False, header=True).value = colName4
            
            strCellexp4 = ws_FactsheetTemplate.range("W44").value
            ws_FactsheetTemplate.range(strCellexp4).options(index=False, header=True).value = exp4

            strCellname4 = ws_FactsheetTemplate.range("W45").value                    
            ws_FactsheetTemplate.range(strCellname4).options(index=False, header=True).value = name4
    
        #sum exp      
        strCellSumExp = ws_FactsheetTemplate.range("W52").value                    
        ws_FactsheetTemplate.range(strCellSumExp).options(index=False, header=True).value = exp1 + exp2 + exp3 + exp4 + exp5 + exp6
        
    
    
        # Construct path for pdf file
        current_work_dir = o.getcwd()
        
        pdf_file_name = strISIN + "_doc"
        pdf_path = Path(current_work_dir, pdf_file_name)
    
        # Save excel workbook as pdf and showing it
        ws_FactsheetTemplate.to_pdf(path=pdf_path)
                
    #show path where PDFs saved 
    ws_FactsheetTemplate.range("V20").options(index=False, header=True).value = current_work_dir


    




if __name__ == "__main__":
    xw.Book("TaxonomyFactsheetGeneration.xlsm").set_mock_caller()
    main()

