Step 1.
Scheduled Trigger-Recurrence
Code View:
 
Step 2.
Obtain Month Day
Inputs: [Link](https://github.com/Sodesu/Data-Analytics-Engineering/blob/80f617f61bfab501df95a887115710e6a878f2ed/Microsoft%20Office%20Apps/Power%20Automate/Obtain_Month_Day.json)
 

Step 3.
Day of Week
Inputs: dayOfWeek(utcNow())
 
Step 4.
Compose FileName
Inputs: concat('CRM41_', formatDateTime(utcNow(), 'dd'), '.xlsx;)
 
Step 5.
Compose Folder
Inputs: Concat('/FolderName/', outputs('Compose_FileName'))
 
Step 6.
Excel Online (Business): Return Total Rows
File: outputs('Compose_Folder')
 
Step 7.
Init variable: Row Poiner
Type: Integer
Value: 0
 
Step 8.
Init Variable: nextRowPointer
Type: Integer
Value: 0
 
Step 9.
Init Variable: TotalRows
Type: Integer
Body: outputs('Return_Total_Rows')['body']['result']
 
Step 10.
Init Variable: BatchSize
Type: Integer
Value: 500
 
Step 11.
Init Variable: startRowIndex
Type: Integer
Value:0
 
Step 12.
Control: Exclude Run on Thursday Saturday and Sunday
Params:
And:
outputs('Day_of_Week') is not equal to 4
outputs('Day_of_Week') is not equal to 6
 outputs('Day_of_Week') is not equal to 0
 
Step 13.
Condition if False:
*Nothing*
 
Step 14.
Condition if True
Delay Action
Count: 3
Unit: Second
 
Step 15.
Excel Online (Business) Clear Rows:
File: Target Workbook in Sharepoint Folder
Script: ClearRowsFromWB
 
Step 16.
Delay
Count: 5
 
Step 17.
Do (container Begins)
Excel Online (Business) Extract Raw MKT Data
File: outputs('Compose_Folder')
Script: Exract Raw MKT Data
 
Step 18.
ScriptParameters/datastartindex:
variables('RowPointer')
 
Step 19.
ScriptParameters/batchsize:
variables('BatchSize')
 
Step 20.
Compose
Inputs: outputs('Extract_Raw_MKT_Data')?['body']['result']['batch'] or body.result.batch
 
Step 21.
Condition Action
And:
Variables('TotalRows') is equal to null
 
Step 21 (False).
If False:
*Nothing*
 
Step 21 (True).
If True:
outputs('Return_Total_Rows')?['body']['result']
(End Condition)
 
Compose - Total Rows Value
Variables('Total_Rows')
 
Excel Online (Business) Load Raw Data:
Script: LoadRowsFromWB
ScriptParameters/batchinput: outputs('Extract_Raw_MKT_Data')?['body']['result']['batch'] or body.result.batch
 
 
Set Variable:
Name: nextRowPointer
Value: add(variables('BatchSize'), variables('RowPointer'))
 
Set Variable 2:
Name: RowPointer
Variables('nextRowPointer')
 
 
Loop Until variables('RowPointer') is greater or equal to variables('TotalRows')
(Do container Ends)
 
Step 22.
Internal Action Condition - If Do Until Loop failed or timed out:
Send an email (V2)
Subject: outputs('Obtain_Month_Day')
Body: Good morning all,
Unfortunately the sharepoint file failed to update either because of a timeout error, or an unforeseen edge case error that didn't appear during the testing of this mechanism.
Please use the File from the most recent update or the original data from two days ago.
Many apologies for the inconvenience caused by this failed update.
Best,
Kunmi
Note - This is an automated message.
 
Step 23.
Internal Action Condition - If Send an email (V2) is skipped:
Excel Online (Business) Refresh Sharepoint File
File: /FolderName/FileName.xlsx
Script: RefreshData
 
Step 24.
Send an Email (V2)
Subject: Delivery Success outputs('Obtain_Month_Day')
Body: Good Morning All,
The Update today was successful.
Please find the updated file within the sharepoint via:
Sharepoint File link
Best,
Kunmi
Note - This is an automated message
 
Step 25.
Internal Action Condition - If Send an email (V2) above is skipped:
Subject: Delivery: outputs('Obtain_Month_Day')
Body: Good morning all,
Unfortunately, an unforseen edge case error occurred resulting in the failure of this mornings delivery.
Please use the File from the most recent update or the original data from two days ago.
Many apologies for the inconvenience caused by this failed update.
Best,
Kunmi
Note - This is an automated message.
