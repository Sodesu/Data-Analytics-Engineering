<center> **Action 1.** <br>
Scheduled Trigger-Recurrence<br>
Code View: [Link]
 
**Action 2.** <br>
Obtain Month Day<br>
Inputs: [Link](https://github.com/Sodesu/Data-Analytics-Engineering/blob/80f617f61bfab501df95a887115710e6a878f2ed/Microsoft%20Office%20Apps/Power%20Automate/Obtain_Month_Day.json)
 

**Action 3.** <br>
Day of Week<br>
Inputs: dayOfWeek(utcNow())
 
**Action 4.** <br>
Compose FileName<br>
Inputs: concat('CRM41_', formatDateTime(utcNow(), 'dd'), '.xlsx;)
 
**Action 5.** <br>
Compose Folder<br>
Inputs: Concat('/FolderName/', outputs('Compose_FileName'))
 
**Action 6.** <br>
Excel Online (Business): Return Total Rows<br>
File: outputs('Compose_Folder')
 
**Action 7.** <br>
Init variable: Row Poiner<br>
Type: Integer<br>
Value: 0
 
**Action 8.** <br>
Init Variable: nextRowPointer<br>
Type: Integer<br>
Value: 0
 
**Action 9.** <br>
Init Variable: TotalRows<br>
Type: Integer<br>
Body: outputs('Return_Total_Rows')['body']['result']
 
**Action 10.** <br>
Init Variable: BatchSize<br>
Type: Integer<br>
Value: 500
 
**Action 11.** <br>
Init Variable: startRowIndex<br>
Type: Integer<br>
Value:0
 
**Action 12.** <br>
Control: Exclude Run on Thursday Saturday and Sunday<br>
Params:<br>
And:<br>
outputs('Day_of_Week') is not equal to 4<br>
outputs('Day_of_Week') is not equal to 6<br>
 outputs('Day_of_Week') is not equal to 0
 
**Action 13.** <br>
Condition if False:<br>
*Nothing*
 
**Action 14.** <br>
Condition if True<br>
Delay Action<br>
Count: 3<br>
Unit: Second
 
**Action 15.** <br>
Excel Online (Business) Clear Rows:<br>
File: Target Workbook in Sharepoint Folder<br>
Script: ClearRowsFromWB
 
**Action 16.** <br>
Delay<br>
Count: 5
 
**Action 17.** <br>
Do (container Begins)<br>
Excel Online (Business) Extract Raw MKT Data<br>
File: outputs('Compose_Folder')<br>
Script: Exract Raw MKT Data
 
**Action 18.** <br>
ScriptParameters/datastartindex:<br>
variables('RowPointer')
 
**Action 19.** <br>
ScriptParameters/batchsize:<br>
variables('BatchSize')
 
**Action 20.** <br>
Compose
Inputs: outputs('Extract_Raw_MKT_Data')?['body']['result']['batch'] or body.result.batch
 
**Action 21.** <br>
Condition Action<br>
And:<br>
Variables('TotalRows') is equal to null
 
**Action 21 (False).** <br>
If False:<br>
*Nothing*
 
**Action 21 (True).** <br>
If True:<br>
outputs('Return_Total_Rows')?['body']['result']<br>
(End Condition)
 
Compose - Total Rows Value<br>
Variables('Total_Rows')
 
Excel Online (Business) Load Raw Data:<br>
Script: LoadRowsFromWB<br>
ScriptParameters/batchinput: outputs('Extract_Raw_MKT_Data')?['body']['result']['batch'] or body.result.batch
 
 
Set Variable:<br>
Name: nextRowPointer<br>
Value: add(variables('BatchSize'), variables('RowPointer'))
 
Set Variable 2:<br>
Name: RowPointer<br>
Variables('nextRowPointer')
 
 
Loop Until variables('RowPointer') is greater or equal to variables('TotalRows')<br>
(Do container Ends)
 
**Action 22.** <br>
Internal Action Condition - If Do Until Loop failed or timed out:<br>
Send an email (V2)<br>
Subject: outputs('Obtain_Month_Day')<br>
Body: Good morning all,<br>
Unfortunately the sharepoint file failed to update either because of a timeout error, or an unforeseen edge case error that didn't appear during the testing of this mechanism.<br>
Please use the File from the most recent update or the original data from two days ago.<br>
Many apologies for the inconvenience caused by this failed update.<br>
Best,<br>
Kunmi<br>
Note - This is an automated message.
 
**Action 23.** <br>
Internal Action Condition - If Send an email (V2) is skipped:<br>
Excel Online (Business) Refresh Sharepoint File<br>
File: /FolderName/FileName.xlsx<br>
Script: RefreshData
 
**Action 24.** <br>
Send an Email (V2)<br>
Subject: Delivery Success outputs('Obtain_Month_Day')<br>
Body: Good Morning All,<br>
The Update today was successful.<br>
Please find the updated file within the sharepoint via:<br>
Sharepoint File link<br>
Best,<br>
Kunmi<br>
Note - This is an automated message
 
**Action 25.** <br>
Internal Action Condition - If Send an email (V2) above is skipped:<br>
Subject: Delivery: outputs('Obtain_Month_Day')<br>
Body: Good morning all,<br>
Unfortunately, an unforseen edge case error occurred resulting in the failure of this mornings delivery.<br>
Please use the File from the most recent update or the original data from two days ago.<br>
Many apologies for the inconvenience caused by this failed update.<br>
Best,<br>
Kunmi<br>
Note - This is an automated message.</center>
