import pandas as pd
import numpy as np

    
def calculateRatings(quantile, step, metricColName, dfMetrics):
    
    
    numLevels = ((1/step) * quantile) + (1-(1/step)) 

    lstConditions = []
    lstChoice = []    

    # for no override ranking
    for level in range(int(numLevels)):

        if level == numLevels - 1:
            strConditionList = eval('np.ceil(dfMetrics.' + metricColName + '.rank(pct=True).mul(' + str(numLevels) + ')) == ' + str(numLevels-level))
        else:
            strConditionList = eval('np.ceil(dfMetrics.' + metricColName + '.rank(pct=True).mul(' + str(numLevels) + ')) == ' + str(numLevels-level))

        lstConditions.append(strConditionList)
            
    for c in range(int(numLevels)):
        if c == 0:
            value = quantile
        else:
            value = value - step     
        lstChoice.append(value)


    dfMetrics[metricColName + '_Rating'] = (
        np.select(
            condlist=np.array(lstConditions),
            choicelist=np.array(lstChoice), 
            default = np.nan))


    return dfMetrics


def calculateThresholds(quantile, step, dfMetrics, ColNum):

    dfThresholds = pd.DataFrame()
    
    lstLowerBound = []
    lstUpperBound = []
    lstRating = []
    
    lstLowerBound.append(0)
    
    numLevels = ((1/step) * quantile) + (1-(1/step))     
    
    for Level in range(int(numLevels)+1):
        
        if Level == 0:
            continue
        
        df = dfMetrics.quantile(Level/numLevels)
        
        if Level == numLevels:
            lstUpperBound.append(df.iloc[ColNum])
            continue
        else:
            lstLowerBound.append(df.iloc[ColNum])    
            lstUpperBound.append(df.iloc[ColNum])
        
    
    for c in range (int(numLevels)):
        if c == 0:
            value = 1
        else:
            value = value + step     
        lstRating.append(value)
    
    lstHeaders = dfMetrics.columns.tolist()
    colheader = lstHeaders[ColNum]
    
    dfThresholds['Rating ' + colheader ] = lstRating
    dfThresholds['LowerBound'] = lstLowerBound
    dfThresholds['UpperBound'] = lstUpperBound    
    
    return dfThresholds      


def calculateManualThresholdRatings(dfThresholds,dfMetrics):
# returns dfmetrics with new ratings cols calculated based on dfThresholds
# dfThresholds should contain 3 cols for each metric    

    cols = len(dfThresholds.axes[1])
    
    lstLowerBounds = []
    lstUpperBounds = []
    
    
    for i in range(cols):
        
        if i == 0:
            continue
        if i % 3 == 0:
            continue
        if i == 1:
            lstLowerBounds.append(dfThresholds[i])
            continue
        if i == 2:
            lstUpperBounds.append(dfThresholds[i])
            continue
        if i % 3 == 1:
            lstLowerBounds.append(dfThresholds[i])
            continue
        if i % 3 == 2:
            lstUpperBounds.append(dfThresholds[i])
            continue
    
    lstRatings = dfThresholds[0]
    
    
    colsMetrics = len(dfMetrics.axes[1])
    
    ratingsLen = len(lstRatings)
    
    for c in range(colsMetrics):
        strColname = 'Rating ' + str(c)
        dfMetrics[strColname] = np.nan
        strMetricsColName = dfMetrics.columns[c]
    
        for r in range(ratingsLen):
    
            dfMetrics[strColname] = np.where(((dfMetrics[strMetricsColName] > lstLowerBounds[c][r]) & (dfMetrics[strMetricsColName] <= lstUpperBounds[c][r])), lstRatings[r], dfMetrics[strColname])
        


    return dfMetrics
