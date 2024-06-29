import sys
sys.path.append('C:\\Users\\Custom py Modules')
import autofit_module as ac

import os
import pandas as pd
from pyxlsb import open_workbook

# columns = [ 'YEARMONTH',	 'ID',	 'Name',	 'CC',	 'RC',	 'Assignment',	 'Start Date',	 'Leaving Date',	 'FTE Absences',	 'Position Level',	 
# 'Individual Target',	 'Individual Self Target',	 'Consolidation',	 'Consolidation Output',	 'Marketing Dist',	 'Gross Total',	 'MKT XVENS',	 'XVENS Self Param1',	 
# 'XVENS Self Param2',	 'Net Total Output',	 'Mkt Total',	 'Conversion',	 'Self Total',	 'Self Param1',	 'Self Param2',	 'Product Obtained',	 'EVs',	 'XVEN DIR DEDUCTION',	 
# 'Total Output Bonus',	 'MKT Conversion',	 'Self Param1 Income',	 'Self Param2 Income',	 'P&P',	 'TP Challenge',	 'Early Achievement',	 'Excess Achievement',	 'MGM Referrals',	 
# 'Personal Vehicle Stipend',	 'Total Income',	 'Threshold Target',	 'Total Real Income',	 'ACRRUAL ADJ',	 'CAPEX']


columns = ["YEARMONTH", "ID", "Position Level", "Total Real Income"]

sheets = ["Sheet2", "Sheet3", "Sheet4a", "Sheet4b", "Sheet5a", "Sheet5b", "Sheet6a", "Sheet6b"]

path_to_workbooks = "C:\\Users\\Python\\Q12024"

months = ["Jan", "Feb", "Mar"]


dataframes = []

for month in months:
    file_name = f'Origin_{month}.xlsb'
    file_path = os.path.join(path_to_workbooks, file_name)
    
    if os.path.exists(file_path):
        print(f"processing workbook: {file_name}")
        with open_workbook(file_path) as wb:
            for sheet in sheets:
                if sheet in wb.sheets:
                    print(f"processing worksheet: {sheet}")
                    df = pd.read_excel(file_path, sheet_name=sheet, engine='pyxlsb')
                    
                    if df.empty:
                        print(f"No data found: {sheet}")
                        
                    else:
                
                        for col in columns:
                            if col not in df.columns:
                                df[col] = pd.NA
                                
                                
                        df = df[columns]
                        dataframes.append(df)
            else:
                print(f"Sheet not found: {sheet}")
                dataframes.append(pd.DataFrame(columns=columns))
                
    else:
        print(f"File not found: {file_path}")
        
if not dataframes:
    print("No Dataframes were created")
    
else:
    print(f"{len(dataframes)} dataframes created")
    
    

combined_df = pd.concat(dataframes, ignore_index=True)

combined_df = combined_df.dropna(how='all')

combined_df['Total Output'] = combined_df['Net Total Output'].combine_first(combined_df['Gross Total Output'])

combined_df.drop(['Gross Total Output','Net Total Output'], axis=1, inplace=True)


combined_df = combined_df[combined_df["Position Level"] != "Variable1"]


outputfile = 'Accumulated_Q1.xlsx'
combined_df.to_excel(outputfile, index=False, engine='openpyxl', sheet_name='Sheet1')

ac.autofit_columns(outputfile, "Sheet1")


## Openpyxl Alternative
# with pd.ExcelWriter(outputfile, engine='xlsxwriter') as writer:
#     combined_df.to_excel(writer, index=False)
    
# writer = pd.ExcelWriter(outputfile, engine='xlsxwriter')     
# combined_df.to_excel(writer, index=False)

