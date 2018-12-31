'------------------------------------------------------------------
'   Description   	  :      Determines whether the most recent lot resulted in an error by means of checking the current total number of notifications against the number before the lot.
'								In this case, no notifications are expected, since the lot entered should be a fresh lot.
'   Project           :      Uniseries Duplicate Check In Lot Cancel Decom
'   Author            :      Paul F.
'   © 2018   Systech International.  All rights reserved
'------------------------------------------------------------------
'   
'   Prologue:
'   - Lot has been processed by server
'   
'   Epilogue:
'   - Error results have been held against the expected outcomes

Window("Menu").WinButton("End Lot").Click @@ hightlight id_;_5835256_;_script infofile_;_ZIP::ssf3.xml_;_
Dialog("Lot Control").WinButton("Yes").Click @@ hightlight id_;_4721034_;_script infofile_;_ZIP::ssf4.xml_;_


'waiting until the LotIsOpen tag is false
Dim iCount: iCount = 0
While ReadTag("InLot-A") = 1 AND iCount < 30
	iCount = iCount + 1
	wait(1)
Wend

dtStartTime = Now()
Dim arrOutput()
foundNotifications = False
iCount = 0

While (NOT foundNotifications) AND iCount < 6
	wait(10)
	iCount = iCount + 1
	ReDim arrOutput(-1)
	ExecuteSQL GetConnectionString, "Guardian.usp_GetMsgQueue", Array(null, 1), Array("State"), arrOutput
	newCount = UBound(arrOutput) + 1
	
	print "old count=" & DataTable.Value("CurrentLogCount",dtglobalsheet)
	print "new count=" & newCount
	
	If CInt(newCount) > CInt(DataTable.Value("CurrentLogCount",dtglobalsheet)) Then
		foundNotifications = True
	End If
Wend




'Test Case 6 - Expectation: found two new notifications
If foundNotifications Then  'got more rows
	print "log TRUE"
	LogResult Environment("Results_File"), True, dtStartTime, Now(), "FT_Notifications", Null, "Verify notifications were successfully generated", "Notifications were generated"
Else
	print "log FALSE"
	LogResult Environment("Results_File"), False, dtStartTime, Now(), "FT_Notifications", Null, "Verify notifications were successfully generated", "Notifications were not generated"
End If

DataTable.Value("CurrentLogCount",dtglobalsheet) = newCount


'Test Case 7 - Expectation: No duplicates in lot, evidenced by No Duplicate Quarantine window
If VbWindow("frmDuplicateAction").Exist(5) Then	
	reporter.ReportEvent micFail, "Single Item Decom Use Case", "Lot ended with duplicates and presents a Duplicate Dialog"
	LogResult Environment("Results_File"), False, dStartTime, Now(), "Single Item Decom Use Case", "N/A", "End a lot that has duplicates", "Lot ends with a Duplicate Dialog"
	'Clicks selection using hot keys
	ExitTest
Else
	LogResult Environment("Results_File"), True, dStartTime, Now(), "Single Item Decom Use Case", "N/A", "End a lot that has duplicates", "Lot ends without a Duplicate Dialog"
End If


Window("Menu").WinButton("Products").Click


