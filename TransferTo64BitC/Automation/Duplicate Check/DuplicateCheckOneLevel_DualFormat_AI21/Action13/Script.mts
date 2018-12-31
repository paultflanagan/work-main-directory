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
Call boolFunc_EndLot()


If VbWindow("frmDuplicateAction").Exist Then
	reporter.ReportEvent micFail, "Dual Format Use Case", "Lot ended with duplicates and presents a Duplicate Dialog"
	LogResult Environment("Results_File"), False, dStartTime, Now(), "Dual Format Use Case", "UNSS-4599", "End a lot that has duplicates", "Lot ends with a Duplicate Dialog"
Else
	reporter.ReportEvent micPass, "3 Level End Lot_Override Quarantine", "First Lot ends as expected without duplicates"
	LogResult Environment("Results_File"), True, dStartTime, Now(), "Dual Format Use Case", "UNSS-4599", "End the first lot", "Lot ends without a Duplicate Dialog"
	'Return to lot control screen
	Window("Menu").WinButton("Lot Control").Click
End If

