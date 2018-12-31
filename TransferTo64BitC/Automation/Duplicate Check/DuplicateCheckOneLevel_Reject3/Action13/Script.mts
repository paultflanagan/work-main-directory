'------------------------------------------------------------------
'   Description   	  :      Determines whether the most recent lot resulted in an error by means of checking the current total number of notifications against the number before the lot.
'								In this case, notifications are expected, since the lot entered should share item data with the previous lot.
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

'confirm no notifications sent
'waiting until the LotIsOpen tag is false
Dim iCount: iCount = 0
While ReadTag("InLot-A") = 1 AND iCount < 30
	iCount = iCount + 1
	wait(1)
Wend

dtStartTime = Now()
ReDim arrOutput(-1)
ExecuteSQL GetConnectionString, "Guardian.usp_GetMsgQueue", Array(null, 1), Array("State"), arrOutput
newCount = UBound(arrOutput) + 1

print "arrOutput contents: " & arrOutput
print "old count=" & DataTable.Value("CurrentLogCount",dtglobalsheet)
print "new count=" & newCount


'Test Case 17 - expecting 2(?) notifications
If CInt(newCount) = CInt(DataTable.Value("CurrentLogCount",dtglobalsheet)) Then     'got same number of rows
	LogResult Environment("Results_File"), True, dtStartTime, Now(), "FT_Notifications", Null, "Verify notification were successful", "No notifications generated"
Else
	LogResult Environment("Results_File"), False, dtStartTime, Now(), "FT_Notifications", Null, "Verify notification were successful", "Notifications were generated"
End If 

'Test Case 18
If VbWindow("frmDuplicateAction").Exist Then
	reporter.ReportEvent micPass, "Single Item Decom Use Case", "Lot ended with duplicates and presents a Duplicate Dialog"
	LogResult Environment("Results_File"), True, dStartTime, Now(), "Single Item Decom Use Case", "N/A", "End a lot that has duplicates", "Lot ends with a Duplicate Dialog"
	'Clicks selection using hot keys
	VbWindow("frmDuplicateAction").VbButton("Cancel").Click
Else
	reporter.ReportEvent micFail, "Single Item Decom Use Case", "Lot ends as unexpected without duplicates"
	LogResult Environment("Results_File"), False, dStartTime, Now(), "Single Item Decom Use Case", "N/A", "End a lot that has duplicates", "Lot ends without a Duplicate Dialog"
End If

'Go to IPS Station screen and Rework
Window("Menu").WinButton("IPS Test Screen").Click
Window("Desktop").WinButton("Re-work").Click
Window("Desktop").WinButton("Decom Item").Click
SwfWindow("SPTReworkOp.exe").SwfEdit("txtScan").Set "(01)11880010000025(21)0000106024"
SwfWindow("SPTReworkOp.exe").SwfButton("Go").Click
'SwfWindow("SPTReworkOp.exe").SwfButton("Yes").Click
SwfWindow("SPTReworkOp.exe").SwfEdit("txtScan").Set "(01)11880010000025(21)0000112267"
SwfWindow("SPTReworkOp.exe").SwfButton("Go").Click
'SwfWindow("SPTReworkOp.exe").SwfButton("Yes").Click

Window("Menu").WinButton("IPS Test Screen").Click
Window("Menu").WinButton("Lot Control").Click


'End the lot by quarantine
Window("Menu").WinButton("End Lot").Click @@ hightlight id_;_5835256_;_script infofile_;_ZIP::ssf3.xml_;_
Dialog("Lot Control").WinButton("Yes").Click

VbWindow("frmDuplicateAction").VbButton("Override the Quarantine").Type micAltDwn + "O" + micAltUp 
Call SecondSig("engineer", "xyzzy")


Window("Menu").WinButton("Products").Click


'script to check for resent notifications
Wait(30)
dtStartTime = Now()
ReDim arrOutput(-1)
ExecuteSQL GetConnectionString, "Guardian.usp_GetMsgQueue", Array(null, 1), Array("State"), arrOutput
newCount = UBound(arrOutput) + 1

print "arrOutput contents: " & arrOutput
print "old count=" & DataTable.Value("CurrentLogCount",dtglobalsheet)
print "new count=" & newCount

'Test Case 20 - expecting 4 notifications
If CInt(newCount) > CInt(DataTable.Value("CurrentLogCount",dtglobalsheet)) Then     'found more rows
	LogResult Environment("Results_File"), True, dtStartTime, Now(), "FT_Notifications", Null, "Verify notification were successful", "Notifications were generated"
Else
	LogResult Environment("Results_File"), False, dtStartTime, Now(), "FT_Notifications", Null, "Verify notification were successful", "Notifications were not generated"
End If


'Preparing baseline_target.txt
'Target = baseline_postlot.txt
Dim FSO, TargetFile, sTargetContents
Set FSO = CreateObject("Scripting.FileSystemObject")
Set TargetFile = FSO.OpenTextFile("C:\baseline_postlot.txt", 1)			'Open file for reading
sTargetContents = TargetFile.ReadAll
TargetFile.Close

Set TargetFile = FSO.OpenTextFile("C:\baseline_target.txt", 2, True)	'Open file for writing, create file if it does not exist
TargetFile.Write(sTargetContents)
TargetFile.Close

