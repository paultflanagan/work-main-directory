'------------------------------------------------------------------
'   Description   	  :      Determines whether the most recent lot resulted in an error by means of checking the Duplicate Check inbox for error notifications
'								In this case, no emails are expected, since the problematic products have been excluded.
'   Project           :      Uniseries Duplicate Check Dual Format
'   Author            :      Paul F.
'   © 2018   Systech International.  All rights reserved
'------------------------------------------------------------------
'   
'   Prologue:
'   - Lot has been processed by server and attempted to have been closed.
'   
'   Epilogue:
'   - All server-sent email notifications, if any, have been processed and the results have been held against the expected outcomes

'Run batch file which runs sql script to check for existence of Email triggers
Dim bEmailFound : bEmailFound = False
Dim count : count = 0

Set Wshell = CreateObject("wscript.shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

'Grabbing a stored no-email-found result for comparison
Set TextFile = FSO.OpenTextFile("C:\EmailCheckEmptyBaseline.txt")
EmptyResult = TextFile.ReadAll
TextFile.Close

'Poll for up to 60 seconds, checking to see if SQL detects an email trigger.
While (Not(bEmailFound) AND (count < 12))
	wait(5)
	count = count + 1
	Wshell.Run "cmd /c CD /d C:\ & EmailCheckAll.bat", 1, true
	Set TextFile = FSO.OpenTextFile("C:\EmailCheckAllResults.txt")
	EntireString = TextFile.ReadAll
	TextFile.Close
	If StrComp(EntireString, EmptyResult) <> 0 Then
		bEmailFound = True
	End If
Wend

'Clear out the Email Table for next test.
Wshell.Run "cmd /c CD /d C:\ & EmailTableClear.bat", 1, true

Set TextFile = nothing
Set FSO = nothing
Set Wshell = nothing

'If an email is found, a trigger was activated when unintended, so log a test failure
If bEmailFound Then
    LogResult Environment("Results_File"), False, dtStartTime, Now(), "Email Validation", "N/A", "Verify Email Generated", "Email is generated unexpectedly"
Else
    LogResult Environment("Results_File"), True, dtStartTime, Now(), "Dataname excluded from duplicate check use case", "UNSS-3159", "Validate email is not received when dataname is excluded from duplicate check.", "Email is not received."
End If
