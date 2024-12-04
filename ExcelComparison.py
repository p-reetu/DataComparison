import pandas as pd
from datetime import datetime
import sys

def createKeyCol(df,columnKeys):
    newKeycol = []
    for row in range(len(df.index)):
        combinationKey = ""
        for key in columnKeys:
            combinationKey = combinationKey + "+" +  str(df[key][row])
        newKeycol.append(combinationKey[1:])
    df["+".join(columnKeys)] = newKeycol
    return df

datetimeNow = datetime.now().strftime("_%m%d%Y_%H%M%S")
isExcelSame = True

file1path = "" #provide file1 path
file2path = "" #provide file2 path
path = "\\".join(file1path.split('\\')[:-1])

columnKeys = [] #provide primary key columns list

# Read Excel files into DataFrames
df1 = pd.read_excel(file1path,engine='openpyxl')
df2 = pd.read_excel(file2path,engine='openpyxl')

rowCount1, rowCount2 = len(df1.index), len(df2.index) 
colCount1, colCount2 = len(df1.columns), len(df2.columns)

if colCount1 != colCount2:
    print("\nBoth Excel files have different number of columns.")
    isExcelSame = False
else:
    if list(df1.columns) != list(df2.columns):
        print("\nBoth Excel files have different headers.")
        isExcelSame = False
    else:
        keyToJoin = "+".join(columnKeys)
        if len(columnKeys) == 1:
            keyToJoin = columnKeys[0]
        else:
            df1 = createKeyCol(df1,columnKeys)
            df2 = createKeyCol(df2,columnKeys)
        #print(keyToJoin)
        df1_ID_col = df1[keyToJoin].tolist()
        df2_ID_col = df2[keyToJoin].tolist()
        #print(df1_ID_col,df2_ID_col)
        df1_extra_rows = df1.loc[~df1[keyToJoin].isin(df2_ID_col)]
        df2_extra_rows = df2.loc[~df2[keyToJoin].isin(df1_ID_col)]
        #print(df1_extra_rows,df2_extra_rows)
        if df1_extra_rows.empty:
            print("\nFirst excel file has no extra rows which are not present in second file.")
        else:
            df1_extra_rows.to_excel(path+"\extrsRowsFile1"+datetimeNow+".xlsx", index=False)

        if df2_extra_rows.empty:
            print("\nSecond excel file has no extra rows which are not present in first file.")
        else:
            df2_extra_rows.to_excel(path+"\extrsRowsFile2"+datetimeNow+".xlsx", index=False)

        joineddf = pd.merge(df1, df2, on=keyToJoin, how='inner')
        joineddf = joineddf.fillna('')
        #print(joineddf)
        headers = list(df1.columns)

        resultdict = {
            keyToJoin:[],
            'ColumnName': [],
            'ValueInFile1': [],
            'ValueInFile2': []
        }
        #print(joineddf)
        for rowIndex in range(len(joineddf.index)):
            for columnHeader in headers:
                if columnHeader == keyToJoin:
                    continue
                if str(joineddf[columnHeader+"_x"][rowIndex]) == "NaT" and str(joineddf[columnHeader+"_y"][rowIndex]) == "NaT":
                    continue
                if joineddf[columnHeader+"_x"][rowIndex] != joineddf[columnHeader+"_y"][rowIndex]:
                    #print(joineddf[columnHeader+"_x"][rowIndex],joineddf[columnHeader+"_y"][rowIndex])
                    resultdict[keyToJoin].append(joineddf[keyToJoin][rowIndex])
                    resultdict['ColumnName'].append(columnHeader)
                    resultdict['ValueInFile1'].append(joineddf[columnHeader+"_x"][rowIndex])
                    resultdict['ValueInFile2'].append(joineddf[columnHeader+"_y"][rowIndex])
        if resultdict[keyToJoin] != []:
            resultdf = pd.DataFrame(resultdict)
            resultdf.to_excel(path+"\Report"+datetimeNow+".xlsx", index=False)
            print("\nTotal number of rows having mismatches: "+str(len(set(resultdict[keyToJoin]))))
            print("\nTotal number of mismatches: "+str(len(resultdict[keyToJoin])))
            isExcelSame = False

print("*"+str(isExcelSame))
