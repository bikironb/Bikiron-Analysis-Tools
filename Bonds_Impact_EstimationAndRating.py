import xlwings as xw
import pandas as pd 
import numpy as np
import msp_functions as f
import warnings
warnings.filterwarnings("ignore")

def main(params):
    wb = xw.Book.caller()

# if ratings = 1 only calculate ratings - for override values
    if params == 1:
        wsScore_Thresholds = wb.sheets['Rating_Thresholds']
        ws_Metrics  = wb.sheets['Estimation&Rating']
        ws_Output  = wb.sheets['Estimation&Rating']        
        
        last_row_MetricData = wb.sheets['Estimation&Rating'].range(5,1).end('down').row
        
        dataThresholds = wsScore_Thresholds.range((39,2),(55,37)).valuemain
        dataMetrics = ws_Metrics.range((5,7),(last_row_MetricData,18)).value
        
        dfThresholds = pd.DataFrame(dataThresholds)                        
        dfMetrics = pd.DataFrame(dataMetrics)  
        
        dfRatings = f.calculateManualThresholdRatings(dfThresholds,dfMetrics)
        
        
        dfRatings.drop(dfRatings.iloc[:, 0:12], inplace = True, axis = 1) 
        ws_Output.range("AG5").options(index=False, header=False).value = dfRatings     
        
        
        return
        
    wb.sheets['Estimation&Rating'].range('A5:XFD100000').clear()
        
#create datasets

    last_row_IDs = wb.sheets['Master Data'].range(5,1).end('down').row

    last_row_MetricData = wb.sheets['Input'].range(6,1).end('down').row

# search headers to identify data cols    
# create dataframe with all cols in row 4     
    DataMasterDataColHeaders = wb.sheets['Master Data'].range((4,1),(4,200)).value
    dfMasterDataHeaders = pd.DataFrame(DataMasterDataColHeaders)
    for row in range(200):
        if dfMasterDataHeaders[0][row] == "ISIN":
            intISINCol = row
        elif dfMasterDataHeaders[0][row] == "Issue Date":
            intIssueDateCol = row
        elif dfMasterDataHeaders[0][row] == "Labelling":
            intLabellingCol = row
        elif dfMasterDataHeaders[0][row] == "Bond Name":
            intBondNameCol = row
        elif dfMasterDataHeaders[0][row] == "Industry":
            intIndustryCol = row
        elif dfMasterDataHeaders[0][row] == "ISSUER ID (REFINITIV)":
            intISSUERIDCol = row
        elif dfMasterDataHeaders[0][row] == "ULTIMATE ISSUER ID (REFINITIV)":
            intULTIMATEISSUERIDCol = row
        elif dfMasterDataHeaders[0][row] == "MSP Details of the use of proceeds":
            intMSPDetailsUseProceeds = row
            
#industry, issuer id, ultimateid
    

#data sets
    DataIDs = wb.sheets['Master Data'].range((5,1),(last_row_IDs,200)).value  
    DataImpact = wb.sheets['Input'].range((6,31),(last_row_MetricData,51)).value
    DataImpactISINs = wb.sheets['Input'].range((6,2),(last_row_MetricData,2)).value
    DataImpactReportDate = wb.sheets['Input'].range((6,207),(last_row_MetricData,207)).value

#create dataframes
    dfMasterData = pd.DataFrame(DataIDs)
    dfIDs = dfMasterData[[intISINCol,intBondNameCol,intLabellingCol]]
    dfUltimate = dfMasterData[intULTIMATEISSUERIDCol]
    dfIssuerID = dfMasterData[intISSUERIDCol]
    dfIndustry = dfMasterData[intIndustryCol]

    dfImpact = pd.DataFrame(DataImpact)
    dfImpactISINs = pd.DataFrame(DataImpactISINs)
    dfReportDate = pd.DataFrame(DataImpactReportDate)
        

        
#drop unwanted cols for IDs
#    dfIDs.drop(dfIDs.iloc[:, 2:18], inplace = True, axis = 1) 

### new headers 'Patients treated', 'Hospital beds added','Farmers supported','Electric cars/trains deployed','EV charging points installed','Railway infrastructure constructed/renovated','New or renovated green buildings','Smart meters installed','Land restored/reforested/certified'

### new headers with no spaces or slashes 'Patients_treated', 'Hospital_beds_added','Farmers_supported','Electric_cars_trains_deployed','EV_charging_points_installed','Railway_infrastructure_constructed_renovated','New_or_renovated_green_buildings','Smart_meters_installed','Land_restored_reforested_certified'

# List impactdata columns ['CO2','RENEWABLE_ADDED','ENERGY_FROM_RENEWABLE','ENERGY_SAVED','WATER_SAVED','WASTE_TREATED','JOBS_CREATED','JOBS_SAVED','SME_FINANCED','PEOPLE_FINANCED','SOCIAL_HOUSING','STUDENTS_SUPPORTED']

# New List impactdata columns ['CO2','RENEWABLE_ADDED','ENERGY_FROM_RENEWABLE','ENERGY_SAVED','WATER_SAVED','WASTE_TREATED','JOBS_CREATED','JOBS_SAVED','SME_FINANCED','PEOPLE_FINANCED','SOCIAL_HOUSING','STUDENTS_SUPPORTED','Patients_treated', 'Hospital_beds_added','Farmers_supported','Electric_cars_trains_deployed','EV_charging_points_installed','Railway_infrastructure_constructed_renovated','New_or_renovated_green_buildings','Smart_meters_installed','Land_restored_reforested_certified']



###

#headers
    dfIDs.columns = ['ISIN','BOND NAME','LABELLING']
    dfUltimate.columns = ['ULTIMATEID']
    dfImpact.columns = ['CO2','RENEWABLE_ADDED','ENERGY_FROM_RENEWABLE','ENERGY_SAVED','WATER_SAVED','WASTE_TREATED', \
                        'JOBS_CREATED','JOBS_SAVED','SME_FINANCED','PEOPLE_FINANCED','SOCIAL_HOUSING','STUDENTS_SUPPORTED','Patients_treated', 'Hospital_beds_added','Farmers_supported','Electric_cars_trains_deployed','EV_charging_points_installed','Railway_infrastructure_constructed_renovated','New_or_renovated_green_buildings','Smart_meters_installed','Land_restored_reforested_certified']
    dfReportDate.columns = ['REPORTDATE'] 
    dfImpactISINs.columns = ['ISIN']
        
#add columns to DFs and format date col
    dfIDs['ISSUERID'] = dfIssuerID
    dfIDs['ULTIMATEID'] = dfUltimate
    dfIDs['INDUSTRY'] = dfIndustry

    frames = [dfImpactISINs,dfImpact]
    dfImpact = pd.concat(frames,axis = 1)
    dfReportDate['REPORTDATE'] = pd.to_datetime(dfReportDate['REPORTDATE'],format='%Y%m%d')    
    frames2 = [dfImpact,dfReportDate]
    dfImpact = pd.concat(frames2,axis = 1)
    


#update energy data cols
    if params == 200:
            ws_config = wb.sheets['Adjustments']
            wp = ws_config.range(7,3).value
            wpw = ws_config.range(8,3).value
            sp = ws_config.range(9,3).value
            spw = ws_config.range(10,3).value
            co2cf = ws_config.range(11,3).value
            
            ws_Metrics  = wb.sheets['Estimation&Rating']

            
            last_row_IDs = wb.sheets['Master Data'].range(5,1).end('down').row

            #DataForCO2 = wb.sheets['Master Data'].range((5,20),(last_row_IDs,20)).value # col "MSP Details of the use of proceeds"

            dfDataForCO2 = dfMasterData[intMSPDetailsUseProceeds] #pd.DataFrame(DataForCO2)
            dfDataForCO2 = dfDataForCO2.to_frame()

            #dfDataForCO2 = dfDataForCO2.rename(columns = {'4' : 'Desc'}, inplace = True)
            #dfDataForCO2 = dfDataForCO2.set_axis(['Desc'],axis=1,inplace = False)
            
            dfDataForCO2.columns = ['Desc']
            
            dfDataForCO2['ISIN'] = dfIDs['ISIN']
                        
            dfImpactCO2Data  = pd.merge(dfImpact, dfDataForCO2, on='ISIN', how='left')
            
#            dfImpactCO2Data['RENEWABLE_ADDED'] = np.where(dfImpactCO2Data.RENEWABLE_ADDED == 0,dfImpactCO2Data.ENERGY_FROM_RENEWABLE/(8760*(wp*wpw + sp*spw)/24),dfImpactCO2Data.RENEWABLE_ADDED) # 8760 - No hours in year
#            dfImpactCO2Data['ENERGY_FROM_RENEWABLE'] = np.where(dfImpactCO2Data.ENERGY_FROM_RENEWABLE == 0,dfImpactCO2Data.RENEWABLE_ADDED*(8760*(wp*wpw + sp*spw)/24),dfImpactCO2Data.ENERGY_FROM_RENEWABLE)

            dfImpactCO2Data['CO2'] = np.where((dfImpactCO2Data.CO2 == 0 & (dfImpactCO2Data.Desc == 'Renewable Energy'))  ,(dfImpactCO2Data.ENERGY_FROM_RENEWABLE/co2cf),dfImpactCO2Data.CO2)
            
            del dfImpactCO2Data['Desc']
            
            dfImpact = dfImpactCO2Data
              
            
#Latest report date dataset
    dfIDsandImpactData = pd.merge(dfIDs, dfImpact, on='ISIN', how='left')
# OK HERE'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    dfIDsandImpactDataLatestDate = dfIDsandImpactData.groupby(['ISIN','LABELLING'], as_index=False)['REPORTDATE'].max()

    #dataframe with data filtered for most recent date
    dfReportDateFinal = pd.merge(dfIDsandImpactData,dfIDsandImpactDataLatestDate, on=['ISIN','LABELLING','REPORTDATE'], how='inner') 


    dfReportDateFinal['ISSUERID'] = dfReportDateFinal.ISSUERID.astype(str) ###### here issuerid becomes same as ultimateid - fixed with right issuerid col #####
    
    dfReportDateFinal['ULTIMATEID'] = dfReportDateFinal.ULTIMATEID.astype(str)
    dfReportDateFinal['ULTIMATEID&LABEL'] = dfReportDateFinal['ULTIMATEID'] + dfReportDateFinal['LABELLING']
    dfReportDateFinal['ISSUERID&LABEL'] = dfReportDateFinal['ISSUERID'] + dfReportDateFinal['LABELLING']

#replace 0s with nans - for calcs
    dfReportDateFinal.replace(0, np.nan, inplace=True)

#MISSING DATA - average of any exisitng ultimate parent data - changed to use ISSUERID
    dfUltimateIDs = dfIDs
    dfUltimateIDs['ULTIMATEID'] = dfUltimate
    dfUltimateIDImpactData = pd.merge(dfUltimateIDs, dfImpact, on='ISIN', how='left')

    dfUltimateIDImpactData['ULTIMATEID'] = dfUltimateIDImpactData.ULTIMATEID.astype(str)
    dfUltimateIDImpactData['ISSUERID'] = dfUltimateIDImpactData.ISSUERID.astype(str)

    dfUltimateIDImpactData['ULTIMATEID&LABEL'] = dfUltimateIDImpactData['ULTIMATEID'] + dfUltimateIDImpactData['LABELLING']
    dfUltimateIDImpactData['ISSUERID&LABEL'] = dfUltimateIDImpactData['ISSUERID'] + dfUltimateIDImpactData['LABELLING']
    

    dfMeans = dfReportDateFinal.groupby(['ULTIMATEID&LABEL'], as_index=False)['CO2','RENEWABLE_ADDED','ENERGY_FROM_RENEWABLE','ENERGY_SAVED', \
                                                                              'WATER_SAVED','WASTE_TREATED','JOBS_CREATED','JOBS_SAVED', \
                                                                            'SME_FINANCED','PEOPLE_FINANCED','SOCIAL_HOUSING','STUDENTS_SUPPORTED','Patients_treated', 'Hospital_beds_added','Farmers_supported','Electric_cars_trains_deployed','EV_charging_points_installed','Railway_infrastructure_constructed_renovated','New_or_renovated_green_buildings','Smart_meters_installed','Land_restored_reforested_certified'].mean()

    dfMeansIssuerID = dfReportDateFinal.groupby(['ISSUERID&LABEL'], as_index=False)['CO2','RENEWABLE_ADDED','ENERGY_FROM_RENEWABLE','ENERGY_SAVED', \
                                                                              'WATER_SAVED','WASTE_TREATED','JOBS_CREATED','JOBS_SAVED', \
                                                                            'SME_FINANCED','PEOPLE_FINANCED','SOCIAL_HOUSING','STUDENTS_SUPPORTED','Patients_treated', 'Hospital_beds_added','Farmers_supported','Electric_cars_trains_deployed','EV_charging_points_installed','Railway_infrastructure_constructed_renovated','New_or_renovated_green_buildings','Smart_meters_installed','Land_restored_reforested_certified'].mean()
        
    dfAllDataFilledEstimatesFinal = pd.merge(dfReportDateFinal, dfMeans, on=['ULTIMATEID&LABEL'], how='left')         

    dfAllDataFilledEstimatesFinal = pd.merge(dfAllDataFilledEstimatesFinal, dfMeansIssuerID, on=['ISSUERID&LABEL'], how='left')  

    dfAllDataNoEstimatesFilled = pd.merge(dfAllDataFilledEstimatesFinal, dfMeansIssuerID, on=['ISSUERID&LABEL'], how='left')        
#################################
    
#drop extra ultimateid and label col
#    dfAllDataFilledEstimatesFinal = dfAllDataFilledEstimatesFinal.drop(dfAllDataFilledEstimatesFinal.columns[17], axis=1)
    
# means data are metrics cols without _x eg CO2     
### populate "Estimates" col - 1 if estimate was used, 0 if data exists, blank if no data exists

# check if all metric values are NULL
# check if any estimates exist
# check for government activity
# if all metrics are NULL and any estimates exist and industry not Gov activity then 1 else 0
    
    dfAllDataNoEstimatesFilled['ESTIMATE'] = np.nan
    
    dfAllDataFilledEstimatesFinal['ESTIMATE'] = np.nan

#if all actuals are null and any of the means are not null then 1, else 2
#then executes this
# if all actuals are null and all means are null then 0

# 1 - all actuals are null and at least one mean has value
# 2 - actuals not null
# 0 - all actuals are null and all means are null

                                                      
    dfAllDataFilledEstimatesFinal.ESTIMATE = np.where((dfAllDataFilledEstimatesFinal.CO2_x.isnull() &                       
                                                      dfAllDataFilledEstimatesFinal.RENEWABLE_ADDED_x.isnull() &  
                                                       dfAllDataFilledEstimatesFinal.ENERGY_FROM_RENEWABLE_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.ENERGY_SAVED_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.WATER_SAVED_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.WASTE_TREATED_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.JOBS_CREATED_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.JOBS_SAVED_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.SME_FINANCED_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.PEOPLE_FINANCED_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.SOCIAL_HOUSING_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.STUDENTS_SUPPORTED_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.Patients_treated_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.Hospital_beds_added_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.Farmers_supported_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.Electric_cars_trains_deployed_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.EV_charging_points_installed_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.Railway_infrastructure_constructed_renovated_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.New_or_renovated_green_buildings_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.Smart_meters_installed_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.Land_restored_reforested_certified_x.isnull()) & \

                                                       ((dfAllDataFilledEstimatesFinal.CO2.notnull() |                                 
                                                        dfAllDataFilledEstimatesFinal.RENEWABLE_ADDED.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.ENERGY_FROM_RENEWABLE.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.ENERGY_SAVED.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.WATER_SAVED.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.WASTE_TREATED.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.JOBS_CREATED.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.JOBS_SAVED.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.SME_FINANCED.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.PEOPLE_FINANCED.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.SOCIAL_HOUSING.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.STUDENTS_SUPPORTED.notnull()) | \
                                                        dfAllDataFilledEstimatesFinal.Patients_treated.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.Hospital_beds_added.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.Farmers_supported.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.Electric_cars_trains_deployed.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.EV_charging_points_installed.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.Railway_infrastructure_constructed_renovated.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.New_or_renovated_green_buildings.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.Smart_meters_installed.notnull() | \
                                                        dfAllDataFilledEstimatesFinal.Land_restored_reforested_certified.notnull()) & \
                                                        
                                                       (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') ,1, 2)  

    dfAllDataFilledEstimatesFinal.ESTIMATE = np.where(dfAllDataFilledEstimatesFinal.CO2_x.isnull() &                       
                                                      dfAllDataFilledEstimatesFinal.RENEWABLE_ADDED_x.isnull() &  
                                                       dfAllDataFilledEstimatesFinal.ENERGY_FROM_RENEWABLE_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.ENERGY_SAVED_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.WATER_SAVED_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.WASTE_TREATED_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.JOBS_CREATED_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.JOBS_SAVED_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.SME_FINANCED_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.PEOPLE_FINANCED_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.SOCIAL_HOUSING_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.STUDENTS_SUPPORTED_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.Patients_treated_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.Hospital_beds_added_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.Farmers_supported_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.Electric_cars_trains_deployed_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.EV_charging_points_installed_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.Railway_infrastructure_constructed_renovated_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.New_or_renovated_green_buildings_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.Smart_meters_installed_x.isnull() & \
                                                       dfAllDataFilledEstimatesFinal.Land_restored_reforested_certified_x.isnull() & \
    
                                                       ((dfAllDataFilledEstimatesFinal.CO2.isnull() &                              
                                                        dfAllDataFilledEstimatesFinal.RENEWABLE_ADDED.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.ENERGY_FROM_RENEWABLE.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.ENERGY_SAVED.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.WATER_SAVED.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.WASTE_TREATED.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.JOBS_CREATED.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.JOBS_SAVED.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.SME_FINANCED.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.PEOPLE_FINANCED.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.SOCIAL_HOUSING.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.STUDENTS_SUPPORTED.isnull()) & \
                                                        dfAllDataFilledEstimatesFinal.Patients_treated.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.Hospital_beds_added.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.Farmers_supported.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.Electric_cars_trains_deployed.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.EV_charging_points_installed.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.Railway_infrastructure_constructed_renovated.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.New_or_renovated_green_buildings.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.Smart_meters_installed.isnull() & \
                                                        dfAllDataFilledEstimatesFinal.Land_restored_reforested_certified.isnull()) & \
                                                        
                                                       (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') ,0, dfAllDataFilledEstimatesFinal.ESTIMATE)  

#populate estimates in dataset
                                               
    dfAllDataFilledEstimatesFinal.CO2_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \
                                                   
                                                   dfAllDataFilledEstimatesFinal.CO2.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.CO2, dfAllDataFilledEstimatesFinal.CO2_x)


    dfAllDataFilledEstimatesFinal.RENEWABLE_ADDED_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \
#                                                  
                                                   dfAllDataFilledEstimatesFinal.RENEWABLE_ADDED.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.RENEWABLE_ADDED, dfAllDataFilledEstimatesFinal.RENEWABLE_ADDED_x)

    
    dfAllDataFilledEstimatesFinal.ENERGY_FROM_RENEWABLE_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \
#                                                  
                                                   dfAllDataFilledEstimatesFinal.ENERGY_FROM_RENEWABLE.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.ENERGY_FROM_RENEWABLE, dfAllDataFilledEstimatesFinal.ENERGY_FROM_RENEWABLE_x)


    dfAllDataFilledEstimatesFinal.ENERGY_SAVED_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \
#                                                  
                                                   dfAllDataFilledEstimatesFinal.ENERGY_SAVED.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.ENERGY_SAVED, dfAllDataFilledEstimatesFinal.ENERGY_SAVED_x)


    dfAllDataFilledEstimatesFinal.WATER_SAVED_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \
#                                                  
                                                   dfAllDataFilledEstimatesFinal.WATER_SAVED.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.WATER_SAVED, dfAllDataFilledEstimatesFinal.WATER_SAVED_x)


    dfAllDataFilledEstimatesFinal.WASTE_TREATED_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \
#                                                  
                                                   dfAllDataFilledEstimatesFinal.WASTE_TREATED.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.WASTE_TREATED, dfAllDataFilledEstimatesFinal.WASTE_TREATED_x)

        
    dfAllDataFilledEstimatesFinal.JOBS_CREATED_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \
#                                                  
                                                   dfAllDataFilledEstimatesFinal.JOBS_CREATED.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.JOBS_CREATED, dfAllDataFilledEstimatesFinal.JOBS_CREATED_x)


    dfAllDataFilledEstimatesFinal.JOBS_SAVED_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \
                                                   
                                                   dfAllDataFilledEstimatesFinal.JOBS_SAVED.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.JOBS_SAVED, dfAllDataFilledEstimatesFinal.JOBS_SAVED_x)
        

    dfAllDataFilledEstimatesFinal.SME_FINANCED_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \
#                                                  
                                                   dfAllDataFilledEstimatesFinal.SME_FINANCED.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.SME_FINANCED, dfAllDataFilledEstimatesFinal.SME_FINANCED_x)


    dfAllDataFilledEstimatesFinal.PEOPLE_FINANCED_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \
#
                                                   dfAllDataFilledEstimatesFinal.PEOPLE_FINANCED.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.PEOPLE_FINANCED, dfAllDataFilledEstimatesFinal.PEOPLE_FINANCED_x)


    dfAllDataFilledEstimatesFinal.SOCIAL_HOUSING_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \
#
                                                   dfAllDataFilledEstimatesFinal.SOCIAL_HOUSING.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.SOCIAL_HOUSING, dfAllDataFilledEstimatesFinal.SOCIAL_HOUSING_x)


    dfAllDataFilledEstimatesFinal.STUDENTS_SUPPORTED_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \

                                                   dfAllDataFilledEstimatesFinal.STUDENTS_SUPPORTED.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.STUDENTS_SUPPORTED, dfAllDataFilledEstimatesFinal.STUDENTS_SUPPORTED_x)


    dfAllDataFilledEstimatesFinal.Patients_treated_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \

                                                   dfAllDataFilledEstimatesFinal.Patients_treated.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.Patients_treated, dfAllDataFilledEstimatesFinal.Patients_treated_x)

    dfAllDataFilledEstimatesFinal.Hospital_beds_added_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \

                                                   dfAllDataFilledEstimatesFinal.Hospital_beds_added.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.Hospital_beds_added, dfAllDataFilledEstimatesFinal.Hospital_beds_added_x)

    dfAllDataFilledEstimatesFinal.Farmers_supported_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \

                                                   dfAllDataFilledEstimatesFinal.Farmers_supported.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.Farmers_supported, dfAllDataFilledEstimatesFinal.Farmers_supported_x)

    dfAllDataFilledEstimatesFinal.Electric_cars_trains_deployed_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \

                                                   dfAllDataFilledEstimatesFinal.Electric_cars_trains_deployed.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.Electric_cars_trains_deployed, dfAllDataFilledEstimatesFinal.Electric_cars_trains_deployed_x)

    dfAllDataFilledEstimatesFinal.EV_charging_points_installed_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \

                                                   dfAllDataFilledEstimatesFinal.EV_charging_points_installed.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.EV_charging_points_installed, dfAllDataFilledEstimatesFinal.EV_charging_points_installed_x)

    dfAllDataFilledEstimatesFinal.Railway_infrastructure_constructed_renovated_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \

                                                   dfAllDataFilledEstimatesFinal.Railway_infrastructure_constructed_renovated.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.Railway_infrastructure_constructed_renovated, dfAllDataFilledEstimatesFinal.Railway_infrastructure_constructed_renovated_x)

    dfAllDataFilledEstimatesFinal.New_or_renovated_green_buildings_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \

                                                   dfAllDataFilledEstimatesFinal.New_or_renovated_green_buildings.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.New_or_renovated_green_buildings, dfAllDataFilledEstimatesFinal.New_or_renovated_green_buildings_x)

    dfAllDataFilledEstimatesFinal.Smart_meters_installed_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \

                                                   dfAllDataFilledEstimatesFinal.Smart_meters_installed.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.Smart_meters_installed, dfAllDataFilledEstimatesFinal.Smart_meters_installed_x)

    dfAllDataFilledEstimatesFinal.Land_restored_reforested_certified_x = np.where(dfAllDataFilledEstimatesFinal.ESTIMATE == 1 & \

                                                   dfAllDataFilledEstimatesFinal.Land_restored_reforested_certified.notnull() & (dfAllDataFilledEstimatesFinal.INDUSTRY != 'G') , \
                                                   dfAllDataFilledEstimatesFinal.Land_restored_reforested_certified, dfAllDataFilledEstimatesFinal.Land_restored_reforested_certified_x)                                    
# calculate percentiles for each metrics and add to df

# look at np.where to exclude estimate = 1 and .rank function

# do same with df with no estimates
# delete cols for below dataframe, join and fill 
# CO2_x = orignal data filled with estimates, CO2_y = Ultimate issuer id logic, CO2 = IssuerID logic

#######################NEW
                                                        #Patients_treated_x
                                                        #Hospital_beds_added_x
                                                        #Farmers_supported_x
                                                        #Electric_cars_trains_deployed_x.isnull() & \
                                                        #EV_charging_points_installed_x.isnull() & \
                                                        #Railway_infrastructure_constructed_renovated_x.isnull() & \
                                                        #New_or_renovated_green_buildings_x.isnull() & \
                                                        #Smart_meters_installed_x.isnull() & \
                                                        #Land_restored_reforested_certified


    dfAllDataFilledEstimatesFinal['Percentile_CO2'] = dfAllDataFilledEstimatesFinal.CO2_x.rank(pct = True) #,na_option='keep'
    dfAllDataFilledEstimatesFinal['Percentile_RENEWABLE_ADDED'] = dfAllDataFilledEstimatesFinal.RENEWABLE_ADDED_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_ENERGY_FROM_RENEWABLE'] = dfAllDataFilledEstimatesFinal.ENERGY_FROM_RENEWABLE_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_ENERGY_SAVED'] = dfAllDataFilledEstimatesFinal.ENERGY_SAVED_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_WATER_SAVED'] = dfAllDataFilledEstimatesFinal.WATER_SAVED_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_WASTE_TREATED'] = dfAllDataFilledEstimatesFinal.WASTE_TREATED_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_JOBS_CREATED'] = dfAllDataFilledEstimatesFinal.JOBS_CREATED_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_JOBS_SAVED'] = dfAllDataFilledEstimatesFinal.JOBS_SAVED_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_SME_FINANCED'] = dfAllDataFilledEstimatesFinal.SME_FINANCED_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_PEOPLE_FINANCED'] = dfAllDataFilledEstimatesFinal.PEOPLE_FINANCED_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_SOCIAL_HOUSING'] = dfAllDataFilledEstimatesFinal.SOCIAL_HOUSING_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_STUDENTS_SUPPORTED'] = dfAllDataFilledEstimatesFinal.STUDENTS_SUPPORTED_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_Patients_treated'] = dfAllDataFilledEstimatesFinal.Patients_treated_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_Hospital_beds_added'] = dfAllDataFilledEstimatesFinal.Hospital_beds_added_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_Farmers_supported'] = dfAllDataFilledEstimatesFinal.Farmers_supported_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_Electric_cars_trains_deployed'] = dfAllDataFilledEstimatesFinal.Electric_cars_trains_deployed_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_EV_charging_points_installed'] = dfAllDataFilledEstimatesFinal.EV_charging_points_installed_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_Railway_infrastructure_constructed_renovated'] = dfAllDataFilledEstimatesFinal.Railway_infrastructure_constructed_renovated_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_New_or_renovated_green_buildings'] = dfAllDataFilledEstimatesFinal.New_or_renovated_green_buildings_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_Smart_meters_installed'] = dfAllDataFilledEstimatesFinal.Smart_meters_installed_x.rank(pct = True)
    dfAllDataFilledEstimatesFinal['Percentile_Land_restored_reforested_certified'] = dfAllDataFilledEstimatesFinal.Land_restored_reforested_certified_x.rank(pct = True)

#drop cols here ################################### WHY????

#    dfAllDataFilledEstimatesFinal.drop(dfAllDataFilledEstimatesFinal.iloc[:, 19:45], inplace = True, axis = 1) 

# calculate percentiles for non estimates only

    dfReportDateFinal['Percentile_CO2_noEst'] = dfReportDateFinal.CO2.rank(pct = True) #,na_option='keep'
    dfReportDateFinal['Percentile_RENEWABLE_ADDED_noEst'] = dfReportDateFinal.RENEWABLE_ADDED.rank(pct = True)
    dfReportDateFinal['Percentile_ENERGY_FROM_RENEWABLE_noEst'] = dfReportDateFinal.ENERGY_FROM_RENEWABLE.rank(pct = True)
    dfReportDateFinal['Percentile_ENERGY_SAVED_noEst'] = dfReportDateFinal.ENERGY_SAVED.rank(pct = True)
    dfReportDateFinal['Percentile_WATER_SAVED_noEst'] = dfReportDateFinal.WATER_SAVED.rank(pct = True)
    dfReportDateFinal['Percentile_WASTE_TREATED_noEst'] = dfReportDateFinal.WASTE_TREATED.rank(pct = True)
    dfReportDateFinal['Percentile_JOBS_CREATED_noEst'] = dfReportDateFinal.JOBS_CREATED.rank(pct = True)
    dfReportDateFinal['Percentile_JOBS_SAVED_noEst'] = dfReportDateFinal.JOBS_SAVED.rank(pct = True)
    dfReportDateFinal['Percentile_SME_FINANCED_noEst'] = dfReportDateFinal.SME_FINANCED.rank(pct = True)
    dfReportDateFinal['Percentile_PEOPLE_FINANCED_noEst'] = dfReportDateFinal.PEOPLE_FINANCED.rank(pct = True)
    dfReportDateFinal['Percentile_SOCIAL_HOUSING_noEst'] = dfReportDateFinal.SOCIAL_HOUSING.rank(pct = True)
    dfReportDateFinal['Percentile_STUDENTS_SUPPORTED_noEst'] = dfReportDateFinal.STUDENTS_SUPPORTED.rank(pct = True)
    dfReportDateFinal['Percentile_Patients_treated_noEst'] = dfReportDateFinal.Patients_treated.rank(pct = True)
    dfReportDateFinal['Percentile_Hospital_beds_added_noEst'] = dfReportDateFinal.Hospital_beds_added.rank(pct = True)
    dfReportDateFinal['Percentile_Farmers_supported_noEst'] = dfReportDateFinal.Farmers_supported.rank(pct = True)
    dfReportDateFinal['Percentile_Electric_cars_trains_deployed_noEst'] = dfReportDateFinal.Electric_cars_trains_deployed.rank(pct = True)
    dfReportDateFinal['Percentile_EV_charging_points_installed_noEst'] = dfReportDateFinal.EV_charging_points_installed.rank(pct = True)
    dfReportDateFinal['Percentile_Railway_infrastructure_constructed_renovated_noEst'] = dfReportDateFinal.Railway_infrastructure_constructed_renovated.rank(pct = True)
    dfReportDateFinal['Percentile_New_or_renovated_green_buildings_noEst'] = dfReportDateFinal.New_or_renovated_green_buildings.rank(pct = True)
    dfReportDateFinal['Percentile_Smart_meters_installed_noEst'] = dfReportDateFinal.Smart_meters_installed.rank(pct = True)
    dfReportDateFinal['Percentile_Land_restored_reforested_certified_noEst'] = dfReportDateFinal.Land_restored_reforested_certified.rank(pct = True)




# drop cols?????????????????????????
#    dfReportDateFinal.drop(dfReportDateFinal.iloc[:, 1:21], inplace = True, axis = 1) 

# NOW REPLACE BLANKS WHERE ESTIMATES EXIST WITH ESTIMATES PERCENTILE VALUES
    dfFinal = pd.merge(dfAllDataFilledEstimatesFinal, dfReportDateFinal, on=['ISIN'], how='inner')        

    dfFinal.Percentile_CO2_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_CO2, dfFinal.Percentile_CO2_noEst)
    dfFinal.Percentile_RENEWABLE_ADDED_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_RENEWABLE_ADDED, dfFinal.Percentile_RENEWABLE_ADDED_noEst)
    dfFinal.Percentile_ENERGY_FROM_RENEWABLE_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_ENERGY_FROM_RENEWABLE, dfFinal.Percentile_ENERGY_FROM_RENEWABLE_noEst)
    dfFinal.Percentile_ENERGY_SAVED_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_ENERGY_SAVED, dfFinal.Percentile_ENERGY_SAVED_noEst)
    dfFinal.Percentile_WATER_SAVED_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_WATER_SAVED, dfFinal.Percentile_WATER_SAVED_noEst)
    dfFinal.Percentile_WASTE_TREATED_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_WASTE_TREATED, dfFinal.Percentile_WASTE_TREATED_noEst)
    dfFinal.Percentile_JOBS_CREATED_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_JOBS_CREATED, dfFinal.Percentile_JOBS_CREATED_noEst)
    dfFinal.Percentile_JOBS_SAVED_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_JOBS_SAVED, dfFinal.Percentile_JOBS_SAVED_noEst)
    dfFinal.Percentile_SME_FINANCED_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_SME_FINANCED, dfFinal.Percentile_SME_FINANCED_noEst)
    dfFinal.Percentile_PEOPLE_FINANCED_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_PEOPLE_FINANCED, dfFinal.Percentile_PEOPLE_FINANCED_noEst)
    dfFinal.Percentile_SOCIAL_HOUSING_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_SOCIAL_HOUSING, dfFinal.Percentile_SOCIAL_HOUSING_noEst)
    dfFinal.Percentile_STUDENTS_SUPPORTED_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_STUDENTS_SUPPORTED, dfFinal.Percentile_STUDENTS_SUPPORTED_noEst)
    dfFinal.Percentile_Patients_treated_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_Patients_treated, dfFinal.Percentile_Patients_treated_noEst)
    dfFinal.Percentile_Hospital_beds_added_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_Hospital_beds_added, dfFinal.Percentile_Hospital_beds_added_noEst)
    dfFinal.Percentile_Farmers_supported_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_Farmers_supported, dfFinal.Percentile_Farmers_supported_noEst)
    dfFinal.Percentile_Electric_cars_trains_deployed_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_Electric_cars_trains_deployed, dfFinal.Percentile_Electric_cars_trains_deployed_noEst)
    dfFinal.Percentile_EV_charging_points_installed_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_EV_charging_points_installed, dfFinal.Percentile_EV_charging_points_installed_noEst)
    dfFinal.Percentile_Railway_infrastructure_constructed_renovated_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_Railway_infrastructure_constructed_renovated, dfFinal.Percentile_Railway_infrastructure_constructed_renovated_noEst)
    dfFinal.Percentile_New_or_renovated_green_buildings_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_New_or_renovated_green_buildings, dfFinal.Percentile_New_or_renovated_green_buildings_noEst)
    dfFinal.Percentile_Smart_meters_installed_noEst_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_Smart_meters_installed_noEst, dfFinal.Percentile_Smart_meters_installed_noEst)
    dfFinal.Percentile_Land_restored_reforested_certified_noEst = np.where(dfFinal.ESTIMATE==1 , dfFinal.Percentile_Land_restored_reforested_certified, dfFinal.Percentile_Land_restored_reforested_certified_noEst)

#??????????????????????????????????????????    
#    dfFinal.drop(dfFinal.iloc[:, 20:32], inplace = True, axis = 1) 


#RATINGS
    #add step: blank out estimate values in ratings function and create dfCO2rating = f.calc.. then add back dfCO2rating col to dfAllDataFilledEstimatesFinal    

    dfFinal = f.calculateRatings(5, 0.25, 'CO2_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'RENEWABLE_ADDED_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'ENERGY_FROM_RENEWABLE_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'ENERGY_SAVED_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'WATER_SAVED_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'WASTE_TREATED_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'JOBS_CREATED_x', dfAllDataFilledEstimatesFinal)    
    dfFinal = f.calculateRatings(5, 0.25, 'JOBS_SAVED_x', dfAllDataFilledEstimatesFinal)    
    dfFinal = f.calculateRatings(5, 0.25, 'SME_FINANCED_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'PEOPLE_FINANCED_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'SOCIAL_HOUSING_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'STUDENTS_SUPPORTED_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'Patients_treated_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'Hospital_beds_added_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'Farmers_supported_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'Electric_cars_trains_deployed_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'EV_charging_points_installed_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'Railway_infrastructure_constructed_renovated_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'New_or_renovated_green_buildings_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'Smart_meters_installed_x', dfAllDataFilledEstimatesFinal)
    dfFinal = f.calculateRatings(5, 0.25, 'Land_restored_reforested_certified_x', dfAllDataFilledEstimatesFinal)    
    
    # delete unwanted columns in between 'REPORTDATE' and 'ESTIMATE'
    dfFinal.drop(dfFinal.iloc[:, 28:72], inplace = True, axis = 1)  

# calculate ratings for non estimates only

#    dfReportDateFinal = f.calculateRatings(5, 0.25, 'CO2_x', dfReportDateFinal)
#    dfReportDateFinal = f.calculateRatings(5, 0.25, 'RENEWABLE_ADDED_x', dfReportDateFinal)
#    dfReportDateFinal = f.calculateRatings(5, 0.25, 'ENERGY_FROM_RENEWABLE_x', dfReportDateFinal)
#    dfReportDateFinal = f.calculateRatings(5, 0.25, 'ENERGY_SAVED_x', dfReportDateFinal)
#    dfReportDateFinal = f.calculateRatings(5, 0.25, 'WATER_SAVED_x', dfReportDateFinal)
#    dfReportDateFinal = f.calculateRatings(5, 0.25, 'WASTE_TREATED_x', dfReportDateFinal)
#    dfReportDateFinal = f.calculateRatings(5, 0.25, 'JOBS_CREATED_x', dfReportDateFinal)    
#    dfReportDateFinal = f.calculateRatings(5, 0.25, 'JOBS_SAVED_x', dfReportDateFinal)    
#    dfReportDateFinal = f.calculateRatings(5, 0.25, 'SME_FINANCED_x', dfReportDateFinal)
#    dfReportDateFinal = f.calculateRatings(5, 0.25, 'PEOPLE_FINANCED_x', dfReportDateFinal)
#    dfReportDateFinal = f.calculateRatings(5, 0.25, 'SOCIAL_HOUSING_x', dfReportDateFinal)
#    dfReportDateFinal = f.calculateRatings(5, 0.25, 'STUDENTS_SUPPORTED_x', dfReportDateFinal)


# code to clear data - after we know layout 

#Place data
    wsImpactRating = wb.sheets['Estimation&Rating']
    wsImpactRating.range("A5").options(index=False, header=False).value = dfFinal

# datapoints count table################################################################ change this to 27
    DataFinal = wb.sheets['Estimation&Rating'].range((5,7),(last_row_IDs,27)).value  
    dfDataFinal = pd.DataFrame(DataFinal)

    dfDataFinal.columns = ['CO2','RENEWABLE_ADDED','ENERGY_FROM_RENEWABLE','ENERGY_SAVED','WATER_SAVED','WASTE_TREATED', \
                        'JOBS_CREATED','JOBS_SAVED','SME_FINANCED','PEOPLE_FINANCED','SOCIAL_HOUSING','STUDENTS_SUPPORTED','Patients_treated', 'Hospital_beds_added','Farmers_supported','Electric_cars_trains_deployed','EV_charging_points_installed','Railway_infrastructure_constructed_renovated','New_or_renovated_green_buildings','Smart_meters_installed','Land_restored_reforested_certified']

    dfDataFinal['Count Type'] = 1

    dfCountsUnique = dfDataFinal.groupby(['Count Type'], as_index=False)['CO2','RENEWABLE_ADDED','ENERGY_FROM_RENEWABLE','ENERGY_SAVED', \
                                                                              'WATER_SAVED','WASTE_TREATED','JOBS_CREATED','JOBS_SAVED', \
                                                                            'SME_FINANCED','PEOPLE_FINANCED','SOCIAL_HOUSING','STUDENTS_SUPPORTED','Patients_treated', 'Hospital_beds_added','Farmers_supported','Electric_cars_trains_deployed','EV_charging_points_installed','Railway_infrastructure_constructed_renovated','New_or_renovated_green_buildings','Smart_meters_installed','Land_restored_reforested_certified'].nunique()

    dfCounts = dfDataFinal.groupby(['Count Type'], as_index=False)['CO2','RENEWABLE_ADDED','ENERGY_FROM_RENEWABLE','ENERGY_SAVED', \
                                                                              'WATER_SAVED','WASTE_TREATED','JOBS_CREATED','JOBS_SAVED', \
                                                                            'SME_FINANCED','PEOPLE_FINANCED','SOCIAL_HOUSING','STUDENTS_SUPPORTED','Patients_treated', 'Hospital_beds_added','Farmers_supported','Electric_cars_trains_deployed','EV_charging_points_installed','Railway_infrastructure_constructed_renovated','New_or_renovated_green_buildings','Smart_meters_installed','Land_restored_reforested_certified'].count()
    dfCounts.apply(pd.to_numeric)
    dfCountsUnique.apply(pd.to_numeric)
        
    dfPercentUnique = dfCountsUnique.div(dfCounts)
    dfPercentUnique = dfPercentUnique.mul(100)
    dfPercentUnique = dfPercentUnique.round()
    
    dfCountsFinal = pd.concat([dfCountsUnique,dfCounts,dfPercentUnique], ignore_index=True)

    dfCountsFinal.iloc[0, 0] = 'Number of Unique Points'    
    dfCountsFinal.iloc[1, 0] = 'Number of Data Points'    
    dfCountsFinal.iloc[2, 0] = '% Unique Data Points'    

#    dfCountsFinal.style.applymap('font-weight: bold', subset=pd.IndexSlice[dfCountsFinal.index[dfCountsFinal.index==[1]], :])
#    dfCountsFinal.style.applymap('font-weight: bold')

    wsScore_Thresholds = wb.sheets('Rating_Thresholds')
    wsScore_Thresholds.range("B8").options(index=False, header=True).value = dfCountsFinal

# # thresholds table

# # get dataset for each metrics
    df0 = f.calculateThresholds(5,0.25,dfDataFinal,0)
    df1 = f.calculateThresholds(5,0.25,dfDataFinal,1)
    df2 = f.calculateThresholds(5,0.25,dfDataFinal,2)
    df3 = f.calculateThresholds(5,0.25,dfDataFinal,3)
    df4 = f.calculateThresholds(5,0.25,dfDataFinal,4)    
    df5 = f.calculateThresholds(5,0.25,dfDataFinal,5)        
    df6 = f.calculateThresholds(5,0.25,dfDataFinal,6)        
    df7 = f.calculateThresholds(5,0.25,dfDataFinal,7)        
    df8 = f.calculateThresholds(5,0.25,dfDataFinal,8)        
    df9 = f.calculateThresholds(5,0.25,dfDataFinal,9)        
    df10 = f.calculateThresholds(5,0.25,dfDataFinal,10)        
    df11 = f.calculateThresholds(5,0.25,dfDataFinal,11)        
    df12 = f.calculateThresholds(5,0.25,dfDataFinal,12)        
    df13 = f.calculateThresholds(5,0.25,dfDataFinal,13)        
    df14 = f.calculateThresholds(5,0.25,dfDataFinal,14)        
    df15 = f.calculateThresholds(5,0.25,dfDataFinal,15)        
    df16 = f.calculateThresholds(5,0.25,dfDataFinal,16)        
    df17 = f.calculateThresholds(5,0.25,dfDataFinal,17)        
    df18 = f.calculateThresholds(5,0.25,dfDataFinal,18)        
    df19 = f.calculateThresholds(5,0.25,dfDataFinal,19)        
    df20 = f.calculateThresholds(5,0.25,dfDataFinal,20)        



    dfAllThresholds = pd.concat([df0, df1, df2, df3, df4, df5, df6, df7, df8, df9, df10, df11, df12, df13, df14, df15, df16, df17, df18, df19, df20] , axis=1, join='inner')

    wsScore_Thresholds.range("B16").options(index=False, header=True).value = dfAllThresholds 

    wsScore_Thresholds.range("B38").options(index=False, header=True).value = dfAllThresholds 
    
    wsScore_Thresholds.range("B37").options(index=False, header=True).value = "Override Upper/Lower Bound Values"


    if params == 2:
        
        ws_config = wb.sheets['Adjustments']
        wp = ws_config.range(7,3).value
        wpw = ws_config.range(8,3).value
        sp = ws_config.range(9,3).value
        spw = ws_config.range(10,3).value
        co2cf = ws_config.range(11,3).value
                
        ws_Metrics  = wb.sheets['Estimation&Rating']
    
        last_row_IDs = wb.sheets['Master Data'].range(5,1).end('down').row
        
        ws_Metrics  = wb.sheets['Estimation&Rating']
        last_row_IDs = wb.sheets['Estimation&Rating'].range(5,1).end('down').row

        dfUseOfProceeds =  dfMasterData[intMSPDetailsUseProceeds]
        ws_Metrics.range("XFD5").options(index=False, header=False).value = dfUseOfProceeds

        for c in range(5,last_row_IDs):

              valCO2 = ws_Metrics.range(c,7).value
              valEnergyAdded = ws_Metrics.range(c,8).value 
              valEnergyProducedFromEneryAdded = ws_Metrics.range(c,9).value 
              valEstimate = ws_Metrics.range(c,20).value
              valUseProceeds = ws_Metrics.range(c,16384).value
              
              if (not(valEnergyProducedFromEneryAdded is None or valEnergyProducedFromEneryAdded <= 0 ) and (valEnergyAdded is None or valEnergyAdded <= 0)) :

                   xw.Range("H" + str(c)).value = valEnergyProducedFromEneryAdded/(8760*(wp*wpw + sp*spw)/24)
                   xw.Range("H" + str(c)).color = (255, 128, 128)

              if (not(valEnergyAdded is None or valEnergyAdded <= 0) and (valEnergyProducedFromEneryAdded is None or valEnergyProducedFromEneryAdded <= 0)) :
                  
                   xw.Range("I" + str(c)).value = valEnergyAdded*(8760*(wp*wpw + sp*spw)/24)
                   xw.Range("I" + str(c)).color = (255, 128, 128)         

              if (not(valEnergyProducedFromEneryAdded is None or valEnergyProducedFromEneryAdded <= 0) and (valCO2 is None or valCO2 <= 0) and (valUseProceeds == 'Renewable Energy')) :
                  
                   xw.Range("G" + str(c)).value = valEnergyProducedFromEneryAdded/co2cf
                   xw.Range("G" + str(c)).color = (255, 128, 128)         

        for c in range(5,last_row_IDs):

              valCO2 = ws_Metrics.range(c,7).value
              valEnergyAdded = ws_Metrics.range(c,8).value 
              valEnergyProducedFromEneryAdded = ws_Metrics.range(c,9).value 
              valEstimate = ws_Metrics.range(c,20).value
              valUseProceeds = ws_Metrics.range(c,16384).value

              if (not(valEnergyProducedFromEneryAdded is None or valEnergyProducedFromEneryAdded <= 0) and (valCO2 is None or valCO2 <= 0) and (valUseProceeds == 'Renewable Energy')) :
                  
                   xw.Range("G" + str(c)).value = valEnergyProducedFromEneryAdded/co2cf
                   xw.Range("G" + str(c)).color = (255, 128, 128)         


if __name__ == "__main__":
    xw.Book("Bonds_Impact_EstimationAndRating.xlsm").set_mock_caller()
    main()

