'------------------------------------------------------------------
'   Description   	  :      Starts lot for Product
'   Project           :      Uniseries Duplicate Check Dual Format
'   Author            :      Paul F.
'   © 2018   Systech International.  All rights reserved
'------------------------------------------------------------------
'   
'   Prologue:
'   - PIM sheet has been loaded for interaction with loaded data
'   
'   Epilogue:
'   - GetAI21 lot is running


Call boolFunc_SearchAndSelectProduct("DualFormat GetAI21 :")
Call boolFunc_StartLot()

'If there is a loss in server connection, this retries the lot start
'Aborts and ends the UFT application for troubleshooting purposes, in the event of an error
If (VbWindow("frmStatus").VbButton("Abort").Exist(1)) Then
	VbWindow("frmStatus").VbButton("Abort").Click
	ExitTest
ElseIf (VbWindow("frmStatus").VbButton("Retry").Exist(1)) Then
	VbWindow("frmStatus").VbButton("Retry").Click
End If
