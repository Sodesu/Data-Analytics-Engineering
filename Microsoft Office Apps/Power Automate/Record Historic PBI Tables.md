
<div align="center">

</b>Action 1.</b><br>
Scheduled Trigger-Recurrence<br>
Code View:

<br>**Action 2.** <br>
Sharepoint - Delete 2nd Workbook<br>
Site Address* <br>
File ID: 2ndWorkbook.xlsx

**Action 3.** <br>
Sharepoint - Get Original Workbook Data<br>
Site Address* <br>
File ID: 1stWorkbook.xlsx<br>
Infer Content Type: Yes

**Action 4.**
Create New 2M File<br>
Folder Path: Target Directory within Sharepoint<br>
File Name: 2ndWorkbook.xlsx<br>
File Content: body('Get_Original_Workbook_Data')

**Action 5.** <br>
Excel Online (Business) - Delete Rows from Original Workbook<br>
File: Original Workbook Directory within Sharepoint<br>
Script: ClearRowsFromA2

**Action 6.** <br>
PowerBI Action Run a Query Against Dataset<br>
Workspace: Workspace Name<br>
Dataset: PBI Model Name<br>
Query Text: EVALUATE 'Target_Table'<br>
Nulls Included: No

**Action 7.** <br>
Apply to Each<br>
Select an output from Previous Steps: outputs('Run_a_Query_Against_Dataset')?['body/firstTableRows']

**Action 8.** <br>
Excel Online (Business) - Add Row Into Table<br>
File: Original Workbook Directory within Sharepoint<br>
Table: Target Excel Table

**Action 9.** <br>
DateFormat: ISO 8601<br>
For each Column Input:<br>
Items('Apply_to_each')?['PowerBI_Target_Table_Name'[ColumnName]']
</div>
