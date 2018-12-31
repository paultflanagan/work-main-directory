'------------------------------------------------------------------
'   Description   	  :      Starts lot for Product
'   Project           :      Uniseries Duplicate Check In Lot Cancel Decom
'   Author            :      Paul F.
'   © 2018   Systech International.  All rights reserved
'------------------------------------------------------------------
'   
'   Prologue:
'   - PIM sheet has been loaded for interaction with loaded data
'   
'   Epilogue:
'   - Decom Notification Single Item lot is running @@ hightlight id_;_1116692_;_script infofile_;_ZIP::ssf5.xml_;_

If Window("Menu").WinButton("Main").Exist(1) Then
	Window("Menu").WinButton("Main").Click
End If

Call boolFunc_SearchAndSelectProduct("Decom Notification Single Item : ")
Call boolFunc_StartLot()

'If there is a loss in server connection, this retries the lot start
'Aborts and ends the UFT application for troubleshooting purposes, in the event of an error
If (VbWindow("frmStatus").VbButton("Abort").Exist(1)) Then
	VbWindow("frmStatus").VbButton("Abort").Click
	ExitTest
ElseIf (VbWindow("frmStatus").VbButton("Retry").Exist(1)) Then
	VbWindow("frmStatus").VbButton("Retry").Click
End If

'prelot
Dim arrOutput(), intCurrentCount
ReDim arrOutput(-1)
ExecuteSQL GetConnectionString, "Guardian.usp_GetMsgQueue", Array(null, 1), Array("State"), arrOutput
DataTable.GlobalSheet.AddParameter "CurrentLogCount", UBound(arrOutput) + 1
print "current log count = " & UBound(arrOutput) + 1
