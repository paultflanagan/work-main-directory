Window("Menu").WinButton("End Lot").Click @@ hightlight id_;_5835256_;_script infofile_;_ZIP::ssf3.xml_;_
Dialog("Lot Control").WinButton("Yes").Click @@ hightlight id_;_4721034_;_script infofile_;_ZIP::ssf4.xml_;_

If VbWindow("frmDuplicateAction").Exist Then
	
reporter.ReportEvent micPass, "Dual Format Use Case", "Lot ended with duplicates and presents a Duplicate Dialog"
LogResult Environment("Results_File"), True, dStartTime, Now(), "Dual Format Use Case", "N/A", "End a lot that has duplicates", "Lot ends with a Duplicate Dialog"

Else

reporter.ReportEvent micFail, "3 Level End Lot_Override Quarantine", "Lot ends as unexpected without duplicates"
LogResult Environment("Results_File"), False, dStartTime, Now(), "Dual Format Use Case", "N/A", "End a lot that has duplicates", "Lot ends without a Duplicate Dialog"

End If

'Clicks selection using hot keys
VbWindow("frmDuplicateAction").VbButton("Override the Quarantine").Type micAltDwn + "O" + micAltUp 
Call SecondSig("engineer", "xyzzy")

If	Dialog("Second User Entry").Exist Then
reporter.ReportEvent micPass, "Second User Entry", "Second User Entry dialog appears and lot is successfuly quarantined"
LogResult Environment("Results_File"), True, dStartTime, Now(), "Second User Entry dialog for Override Quarantine", "UNSS-3159", "Enter second signature", "Lot is quarantined with second signature"

Else

reporter.ReportEvent micFail, "Second User Entry", "Lot ended as expected without duplicates"
LogResult Environment("Results_File"), False, dStartTime, Now(), "Second User Entry dialog for Override Quarantine ", "UNSS-3159", "Enter second signature", "Second User Entry dialog does not appear"

End If

Window("Menu").WinButton("Products").Click

'Advisor SP
'Dim strSQL : strSQL = ReadFile("C:\Automation\Duplicate Check\SQL\usp_OperationGetSPTList-Rollback.sql")
'ExecuteSQL "Provider=SQLOLEDB;Data Source=ENGADVDEV03;Password=TUser!13;User ID=TipsApp;Initial Catalog=TipsDB;", strSQL, Null, Null, NULL
'LogResult Environment("Results_File"), True, dtStartTime, Now(), "Toggle_SQL", Null, "Reset stored procedures in Advisor", "Stored procedures set"
'    
''Guardian SP
'strSQL = ReadFile("C:\Automation\Duplicate Check\SQL\usp_SPTListReturnCheckSet-Rollback.sql")
'ExecuteSQL GetConnectionString, strSQL, Null, Null, NULL
'LogResult Environment("Results_File"), True, dtStartTime, Now(), "Toggle_SQL", Null, "Reset stored procedures in Guardian", "Stored procedures set"
'    

