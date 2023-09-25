'''
Reads a text formatted Depreciation report to summarize totals by GL account
1) Finds data rows
2) imports into pandas dataframe
3) exports to excel
'''
# This code has been edited to ensure anonymity



# import modules
from pathlib import Path
import pandas as pd
import numpy as np
import os, re

def read_SourceFile(fileName):        # opens file and reads lines
        global fileContent
        readFile = open(fileName)
        fileContent = readFile.readlines()
        readFile.close()
        return fileContent

def extractLines(fileContent):          # searches lines for gl account number, if found adds the line for output
        #fileOutput = '' # initialize empty string

        #creates headers for output files
        fileOutput = 'Account CostCtr Loc   BeginningBal Additions    Depreciation  Adjustments   Retirements   Reclasses     Transfers     Ending_Bal   \n'

        for line in fileContent:          
            
                # Regex search for reserve account for line to import

                textRegex = re.compile(r'\d\d\d\d\d\d\s\s\d\d\d\d\s\s\s\s\d\d\d\d')
                results = textRegex.findall(line)

                # if asset# found write the line to the list
                if len(results) != 0:

                        writeLine = (line)
                        fileOutput = (fileOutput + writeLine)
 

        # write output to temporary fixed width file
        outputName = Path(fileDirectory / 'AssetDepreciation.txt')        
        outPutFile = open(outputName, 'w')
        outPutFile.write(fileOutput)
        outPutFile.close()


def createDataFrame(fileDirectory):
        global depr_df
        # read in temporary fixed width file
        colWidths = (8, 8, 6, 13, 13, 14, 14, 14, 14, 14, 13)
        report_df = pd.read_fwf(fileDirectory / 'AssetDepreciation.txt', widths = colWidths)

        # CLEAN UP- BEGIN
        # create subset of report dataframe
        # include only rows >= 6, and columns account & depreciation
        depr_df = report_df[6:][['Account', 'Depreciation']]

        depr_df.index = np.arange(0,len(depr_df))	                        # reindexes subset dataframe
        depr_df.fillna(0, inplace = True)                                       # set NA values to 0
        depr_df['Depreciation'].replace(',','', regex=True, inplace=True)      # remove commas from numeric fields

        convert_dict = {'Account': int, 'Depreciation': float}
        depr_df = depr_df.astype(convert_dict)                                  # convert columns to numeric fields
        # CLEAN UP- COMPLETE

        results_df = depr_df.groupby(['Account']).sum() # new dataframe with totals by gl account#

        print(results_df)

        results_df.to_excel(fileDirectory / 'depreciation_summary.xlsx') # export summary results to excel
        print('exported')
              

if __name__ == '__main__':

        fileDirectory = Path(Path.home() / 'Documents/Test_Files/')
        fileName = (fileDirectory / 'Depreciation_Report.txt')
        read_SourceFile(fileName)
        extractLines(fileContent)
        createDataFrame(fileDirectory)




