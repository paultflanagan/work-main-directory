' HEADER
'   Description     :      Lot Run
'   Project         :      Shared Tests
'   Author          :      Alex Chepovetsky
'   © 2017   Systech International.  All rights reserved
'------------------------------------------------------------------
'   
'   Prologue:
'   - Ends lot
'   
'   Epilogue:
'   - Not Applicable

'End the lot
Call boolFunc_EndLot() @@ hightlight id_;_4721034_;_script infofile_;_ZIP::ssf4.xml_;_

If VbWindow("frmDuplicateAction").Exist Then
	reporter.ReportEvent micPass, "Dual Format Use Case", "Lot ended with duplicates and presents a Duplicate Dialog"
	LogResult Environment("Results_File"), True, dStartTime, Now(), "Dual Format Use Case", "UNSS-4599", "End the second lot", "Lot ends with a Duplicate Dialog"
Else
	reporter.ReportEvent micFail, "3 Level End Lot_Override Quarantine", "Lot ends as unexpected without duplicates"
	LogResult Environment("Results_File"), False, dStartTime, Now(), "Dual Format Use Case", "UNSS-4599", "End the second lot", "Lot ends without a Duplicate Dialog"
End If


If VbWindow("frmDuplicateAction").Exist Then
	'Clicks selection using hot keys
	VbWindow("frmDuplicateAction").VbButton("Override the Quarantine").Type micAltDwn + "O" + micAltUp
	Window("Screen Manager Diagnostic").Dialog("Second User Entry").Activate	
	Call SecondSig("engineer", "xyzzy")
	Dim iCount: iCount = 0
	'waiting until the LotIsOpen tag is false
	While ReadTag("InLot-A") = 1 AND iCount < 20
		iCount = iCount + 1
		wait(1)
	Wend
	If ReadTag("InLot-A") = 0 Then
		reporter.ReportEvent micPass, "Second User Entry", "Second User Entry dialog appears and lot is successfuly quarantined"
		LogResult Environment("Results_File"), True, dStartTime, Now(), "Second User Entry dialog for Override Quarantine", "UNSS-3159", "Enter second signature", "Lot is quarantined with second signature"
	Else
		reporter.ReportEvent micPass, "Second User Entry", "Second User Entry dialog appears, but lot is unsuccessfuly quarantined"
		LogResult Environment("Results_File"), False, dStartTime, Now(), "Second User Entry dialog for Override Quarantine", "UNSS-3159", "Enter second signature", "Lot quarantine not approved"
	End If
Else
	reporter.ReportEvent micFail, "Second User Entry", "Lot ended unexpectedly without duplicates"
	LogResult Environment("Results_File"), False, dStartTime, Now(), "Second User Entry dialog for Override Quarantine ", "UNSS-3159", "Enter second signature", "Second User Entry dialog does not appear"
End If

Window("Menu").WinButton("Products").Click

'Preparing baseline_target.txt
'Target = baseline_dualformat.txt
Dim FSO, TargetFile, sTargetContents
Set FSO = CreateObject("Scripting.FileSystemObject")
Set TargetFile = FSO.OpenTextFile("C:\baseline_dualformat.txt", 1)		'Open file for reading
sTargetContents = TargetFile.ReadAll
TargetFile.Close

Set TargetFile = FSO.OpenTextFile("C:\baseline_target.txt", 2, True)	'Open file for writing, create file if it does not exist
TargetFile.Write(sTargetContents)
TargetFile.Close
