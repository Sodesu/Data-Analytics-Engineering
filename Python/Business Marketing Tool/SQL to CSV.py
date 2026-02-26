import pyodbc
import pandas as pd
import os
import csv
from datetime import datetime, timedelta
from openpyxl.utils.cell import get_column_letter
import time
import yaml
import sys
import threading
import traceback
import win32com.clParam1Variable2nt as win32

#%% Defined directory paths and variables
timestamp = datetime.now().strftime('%Y%m%d')

json_direct = 'C:\\Users\\Anonymous\\'

BAT_File_Path = 'C:\\Users\\origin.xlsx'

output_file = f'C:\\Users\\output_{timestamp}.csv'

timevariable1 = datetime.now()
timevariable2 = timevariable1 + timedelta(hours=#)
#%%




#%% This is PostgreSQl specific syntax
# conn = pg8000.connect(
    # user='user',
    # password='password',
    # host='host/serve',
    # port=value,
    # DATABASE='database'
# )
#%%


with open(json_direct + 'keys.yaml', 'r') as file:
    config = yaml.safe_load(file)
    
db_config = config['database']


connection_str = (
    'DRIVER={{{driver}}};'
    'SERVER={server};'
    'DATABASE={database};'
    'UID={user};'
    'PWD={password}').format(**db_config)

connection = None
# stop_timer = threading.Event()


# def live_timer():
#     start_time = time.time()
#     while not stop_timer.is_set():
#         elapsed_time = time.time() - start_time
#         sys.stdout.write(f"\rQuery runnning: {elapsed_time:.2f} seconds")
#         sys.stdout.flush()
#         time.sleep(1)
#     elapsed_time = time.time() - start_time
#     print(f"\rQuery completed in {elapsed_time:.2f} seconds")
    
    
    

try:
    connection = pyodbc.connect(connection_str)
    print("Connection success")
    cursor = connection.cursor()    
except Exception as e:
    print("Connection failed")
    print(e)


    
try: 
    #%% Provides all the column names within target_SQL_Table contained in a SELECT statement

    # SELECT TOP 100 Column Name 1, Column Name 2...
    #%%

    #%% Original query takes more than 2 minutes to complete
    # query = """
    # SELECT Column Name 1, Column Name 2, Column Name 3, Column Name 4, Column Name 5,	Column Name 6, Column Name 7, Column Name 8
    # FROM target_SQL_Table
    # WHERE Parameter 1 (not in selected columns) = 'Param1Variable2' AND Parameter 2 (not in selected columns) = 1 AND Column Name 4 = 'Col4Variable5'
    # """
    #%%
    
    SQL_Table_query = """
    SELECT Column Name 1, Column Name 2, Column Name 3, Column Name 4, Column Name 5,	Column Name 6, Column Name 7, Column Name 8
    FROM target_SQL_Table
    WHERE Column Name 4 = 'Col4Variable5' AND Parameter 1 (not in selected columns) = 'Param1Variable2' AND Column Name 7 BETWEEN ? AND ? 
    """
    
    
    params = (timevariable1, timevariable2)
    
    cursor.execute(SQL_Table_query, params)   
    rows = cursor.fetchall()
    columns = [column[0] for column in cursor.description]
    
    SQL_Output_df = pd.DataFrame.from_records(rows, columns=columns)
    
    SQL_Output_df = SQL_Output_df.drop_duplicates()
    
    
    
    #%% Includes logic that queries Param1Variable2 target_SQL_Table#2 for Col1Variables and Col2Variables & filters target_SQL_Table
    # target_SQL_query#2 = """
    # SELECT TOP 100 Column Name 1, Column Name 2
    # FROM target_SQL_Table#2
    # WHERE Column3 = 'Param1Variable2'
    # """
    
    # print("Executing Query...\n")
    
    
    # cursor.execute(target_SQL_query#2)   
    # rows = cursor.fetchall()
    # columns = [column[0] for column in cursor.description]


    
    # target_SQL_Table#2_df = pd.DataFrame.from_records(rows, columns=columns)
    
    
    # filtered_target_SQL_Table = SQL_Output_df[
    #     (SQL_Output_df['Parameter 1 (not in selected columns)'].str.contains("Param1Variable2")) &
    #     (SQL_Output_df['Parameter 2 (not in selected columns)'] == 1) &
    #     (SQL_Output_df['Column Name 4'] == 'Col4Variable5')
    #     ]
    #%%
    
    #%% Merges SQL queried Param1Variable2 dataframes via an inner join on Col1Variables
    # merged_output_df = pd.merge(target_SQL_Table#2_df, filtered_target_SQL_Table, on = "PrimaryKeyCol")
    #%%
    
    #%% To generate column names of the target_SQL_Table table
    # dtarget_SQL_Table_columns = pd.DataFrame(columns, columns=["target_SQL_Table Columns"])
    
    # with pd.ExcelWriter('C:\\Users\\Column_Names.xlsx', engine='openpyxl') as writer:
    #     target_SQL_Table_columns.to_excel(writer, sheet_name='Columns', index=False)
    #%%
    
    
except Exception as e:
    print(f"Error occurred due to: {e}")
    connection.rollback()
    
finally:
    cursor.close()
    connection.close()
    print("Connection closed")

    
if not SQL_Output_df.empty:
    
    SQL_Output_df.to_csv(output_file, index=False)
    print("\nCSV file exported successfully")

    ## Alternative .xlsx formatted output
    # SQL_Output_df.to_excel(output_file, index=False)
    
else:
    print("No data within SQL_Output_df")

