module_directory = 'C:\\Users\\Void\\Custom py Modules'
import sys
sys.path.append(module_directory)
import autofit_module as ac
import os
import pandas as pd
from pyxlsb import open_workbook as open_xlsb
import xlsxwriter
import datetime
from datetime import datetime
from dateutil.relativedelta import relativedelta




def obtain_Primary_Int_Output_month(a):
    current_month = datetime.now()
    Primary_Int_Output_month = current_month - relativedelta(months=a)
    return Primary_Int_Output_month.strftime('%B') # %B is the directive for the full month name




Primary_Int_Outputs_Dir = 'C:\\LocalPath' + obtain_Primary_Int_Output_month(1)

excel_file = Primary_Int_Outputs_Dir + '\\Source\\DataOrigin.xlsb'  


def obtain_data(wb, sheet_name, data_dict, skip_first_row=False):
    
    if sheet_name in wb.sheets:
        with wb.get_sheet(sheet_name) as sheet:
            header_processed=False
            for i, row in enumerate(sheet.rows()):
                row_values = [cell.v for cell in row]
                
                if skip_first_row and i == 0:
                    continue
                
                if not header_processed:
                    header = row_values
                    header_processed =True
                    continue
                
                
                if row_values[0] is None or row_values[0] == '':
                    break
                
                for col_name, value in zip(header, row_values):
                    if col_name in data_dict:
                        data_dict[col_name].append(value)
                            
    return pd.DataFrame(data_dict)
    

with open_xlsb(excel_file) as wb:
    
 
    Sheet1_data = {col: [] for col in ["Sheet1Col1 Integer", "Sheet1Col2 Date", "Sheet1Col3 ID", "Sheet1Col4 ID", "CType", "Sheet1Col6 Integer", "Sheet1Col7", 'Sheet1Col8 Integers']}

    Sheet2_data = {col: [] for col in ["Sheet2Col1 String","Sheet2Col2 String", "Sheet2Col3 String", "Sheet2Col4 String", "Sheet2Col5 String", "Sheet2Col6 String", "Sheet2Col7 String", "Sheet2Col8 Integer"]}
    
    Sheet3_data = {col: [] for col in ["Sheet3Col1 Integer", "Sheet3Col2 String", "Sheet3Col3 String", "Sheet3Col4 String", "CType", "Sheet3Col6 Integer", "Sheet3Col8 Integer", "Primary Int Output", "VariableS Primary Int Output"]}
    



    Sheet1_df = obtain_data(wb, 'Sheet1', Sheet1_data)     
    
    Sheet2_df = obtain_data(wb, 'Sheet2', Sheet2_data)
    
    Sheet3_df = obtain_data(wb, 'Sheet3', Sheet3_data, skip_first_row=True)



Sheet1_filtered_df = Sheet1_df[Sheet1_df['Sheet1Col3 ID'] == 700975] # to obtain the entire Sheet1 data, comment this out 


Sheet1_filtered_df = Sheet1_filtered_df.reset_index(drop=True)


Sheet2_df['Sheet2Col1 String'] = pd.to_numeric(Sheet2_df['Sheet2Col1 String'], errors='coerce') ## required to prevent ID filter from being represented as a string and resolved ValueError: Unable to parse string

filtered_Sheet2_df = Sheet2_df[Sheet2_df['Sheet2Col1 String'] == 700975]

#df.loc[7] = df.loc[7].apply(lambda x: 'CTypeVariable1') ## to replace all the values or
filtered_Sheet2_df['Sheet2Col7 String'] = filtered_Sheet2_df['Sheet2Col7 String'].apply(lambda x: 'CTypeVariable1' if x == 'CTypeVar1' else x) # to replace specific string



full_names = filtered_Sheet2_df['Sheet2Col4 String'].astype(str).str.cat(filtered_Sheet2_df['Sheet2Col5 String'], sep=' ') # Concatenating the first name and last name

filtered_Sheet2_df = filtered_Sheet2_df.drop(columns=['Sheet2Col4 String', 'Sheet2Col5 String'])

filtered_Sheet2_df.insert(3, 'Name', full_names )

adjusted_column_names = ["Sheet1Col1 Integer","Renamed DataFrame Int Col", "Renamed DataFrame Col2 String", "Renamed DataFrame Col3 String", "Renamed DataFrame Col4 String", "CType", "Renamed DataFrame Col5 Int"]

filtered_Sheet2_df.columns = adjusted_column_names

filtered_Sheet2_df = filtered_Sheet2_df.reset_index(drop=True)


#Logic to float one column:
# Sheet3_df['Sheet2Col8 Integer'] = pd.to_numeric(Sheet3_df['Sheet2Col8 Integer'], errors='coerce')
# Sheet3_df['Sheet3Col2 String'] = pd.to_numeric(Sheet3_df['Sheet3Col2 String'], errors='coerce')

# Logic to float multiple columns 
float_columns = ['Sheet2Col8 Integer', 'Sheet3Col2 String']
for column in float_columns:
    Sheet3_df[column] = pd.to_numeric(Sheet3_df[column], errors='coerce')



# Consolidating data
filtered_Sheet3_df = Sheet3_df[pd.notna(Sheet3_df['Sheet3Col2 String']) & (Sheet3_df['Sheet2Col8 Integer'] > 0)]


# Custom Formatting
filtered_Sheet3_df['Sheet3Col2 String'] = 700975
filtered_Sheet3_df.rename(columns={'Sheet3Col2 String': 'Salesman', 'Sheet3Col1 Integer': 'Sheet1Col1 Integer' }, inplace=True)
ins_values = Sheet1_filtered_df['Sheet1Col1 Integer'].unique()
direct_IntOut_data = filtered_Sheet3_df[filtered_Sheet3_df['Sheet1Col1 Integer'].isin(ins_values)]

direct_IntOut_data = direct_IntOut_data.reset_index(drop=True)



output_workbook = Primary_Int_Outputs_Dir + "\\Queries Dir\\Employee_Name.xlsx"



with pd.ExcelWriter(output_workbook, engine='xlsxwriter') as writer:
    Sheet1_filtered_df.to_excel(writer, sheet_name='Sheet1 Data', startrow=1, startcol=0, index_label ='Count')
    
    workbook = writer.book
    worksheet = writer.sheets['Sheet1 Data']
    
    
    
    # merged_named_cell = workbook.add_format({'align': 'center',
    #                                          'bold': True, 
    #                                          'bg_color': '#808080', # Grey background'
    #                                          'font_color': '#FFFFFF'}) # White font applied
    
    
    
    worksheet.merge_range('A1:I1', "Employee_Name" + " " + obtain_Primary_Int_Output_month(1) + " Invoiced Sheet1 Data", workbook.add_format({'align': 'center',
                                             'bold': True, 
                                             'bg_color': '#808080', # Grey background'
                                             'font_color': '#FFFFFF'})) # White font applied
    
    date_format = workbook.add_format({'num_format': 'DD/MM/YYYY'})
    
    column_index = Sheet1_filtered_df.columns.get_loc('Sheet1Col2 Date')
    column_letter = chr(ord('A') + column_index + 1)
    
    worksheet.set_column(f"{column_letter}:{column_letter}", None, date_format)
    
    ## Alternatively
    # for row index in range(3, 3 + len(filtered_df)):
        # worksheet.write_datetime(row_index, column_index + 1, df.iloc[row_index - 3, column_index], date_format)
    
    border_format = workbook.add_format({'border': 1, 'align': 'center'})
    
    worksheet.conditional_format('A2:I' + str(len(Sheet1_filtered_df)+2), {'type': 'no_errors', 'format':border_format})    
    
    

##################### Logic to add the Sheet2 Data to another worksheet within the Workbook ##########################

    filtered_Sheet2_df.to_excel(writer, sheet_name='Sheet2 Data', startrow=1, startcol=0, index_label ='Count')
    
    workbook = writer.book
    worksheet = writer.sheets['Sheet2 Data']
    
    
########## Assigned variable to make code more readable for worksheet.merge_range function ############

    # merged_named_cell = workbook.add_format({'align': 'center',
    #                                          'bold': True, 
    #                                          'bg_color': '#808080', # Grey background'
    #                                          'font_color': '#FFFFFF'}) # White font applied
    
######################################################################################################  
    
    
    worksheet.merge_range('A1:H1', "Employee_Name" + " " + obtain_Primary_Int_Output_month(1) + " Invoiced Sheet2 Data", workbook.add_format({'align': 'center',
                                             'bold': True, 
                                             'bg_color': '#808080', # Grey background'
                                             'font_color': '#FFFFFF'})) # White font applied
    
    
############ Logic to programmatically convert column to number if I didnt already convert ID strings to number format #############
    
    col_num = filtered_Sheet2_df.columns.get_loc('ID') + 1
    for row_num in range(2, len(filtered_Sheet2_df) + 2):
        cell_value = filtered_Sheet2_df.iloc[row_num - 2, col_num - 1]
        cell_value = float(cell_value)
        worksheet.write(row_num, col_num, cell_value * 1)
        
        
      #   worksheet.write_formula(row_num, col_num, f'={cell_loc} * 1') # this creates a circular reference within the column resulting in the 
      #   value of 0, so never use write_formula in attempt to convert the string formatted IDs to numbers
                    
###################################################################################################################################


    border_format = workbook.add_format({'border': 1})
    
    worksheet.conditional_format('A2:H' + str(len(filtered_Sheet2_df)+2), {'type': 'no_errors', 'format':border_format})
    
    
##################### Logic to add the Direct Primary Int Output breakdown Data to another worksheet within the Workbook ##########################

    direct_IntOut_data = direct_IntOut_data.drop('VariableS Primary Int Output', axis=1)
    direct_IntOut_data = direct_IntOut_data.sort_values(by=['Sheet1Col1 Integer', "Sheet3Col6 Integer"], ascending=[True, False])
    direct_IntOut_data.to_excel(writer, sheet_name='Sheet3 Data', startrow=1, startcol=0, index=False)

    
    workbook = writer.book
    worksheet = writer.sheets['Sheet3 Data']



    # merged_named_cell = workbook.add_format({'align': 'center',
    #                                          'bold': True, 
    #                                          'bg_color': '#808080', # Grey background'
    #                                          'font_color': '#FFFFFF'}) # White font applied



    worksheet.merge_range('A1:H1', "Employee_Name"  + " " + obtain_Primary_Int_Output_month(1) + " Invoiced Sheet3 Data", workbook.add_format({'align': 'center',
                                             'bold': True, 
                                             'bg_color': '#808080', # Grey background'
                                             'font_color': '#FFFFFF'})) # White font applied
    

    
    border_format = workbook.add_format({'border': 1, 'align': 'center'})
    
    worksheet.conditional_format('A2:H' + str(len(direct_IntOut_data)+2), {'type': 'no_errors', 'format':border_format})





ac.autofit_columns(output_workbook, 'Sheet1 Data')
ac.autofit_columns(output_workbook, 'Sheet2 Data')
ac.autofit_columns(output_workbook, 'Sheet3 Data')
os.startfile(output_workbook)
