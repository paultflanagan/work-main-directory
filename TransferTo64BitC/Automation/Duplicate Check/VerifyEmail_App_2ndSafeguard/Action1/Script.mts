'------------------------------------------------------------------
'   Description   	  :      Determines whether the most recent lot resulted in an error by means of checking the Duplicate Check inbox for error notifications
'								In this case, no emails are expected, due to the 2nd Safeguard being in place.
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

'Summary: Runs a batch file which runs sql script to check for existence of Email triggers

Dim bEmailFound : bEmailFound = False
Dim count : count = 0

Set Wshell = CreateObject("wscript.shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

'Grabbing no-email-found result for comparison
Set TextFile = FSO.OpenTextFile("C:\EmailCheckEmptyBaseline.txt")
EmptyResult = TextFile.ReadAll
TextFile.Close

'Poll for up to 60 seconds, checking to see if SQL detects an email trigger.
While (Not(bEmailFound) AND (count < 12))
	wait(5)
	count = count + 1
	Wshell.Run "cmd /c CD /d C:\ & EmailCheckReuse.bat", 1, true
	Set TextFile = FSO.OpenTextFile("C:\EmailCheckReuseResults.txt")
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

'If an email is found, then the trigger was activated correctly, so log a test success
If bEmailFound Then
	LogResult Environment("Results_File"), True, dtStartTime, Now(), "Quarantined Dupes use case", "UNSS-3159", "Validate email is received when dupes are quarantined.", "Email received as expected."
Else
	LogResult Environment("Results_File"), False, dtStartTime, Now(), "Email Validation", "UNSS-3159", "Validate the email", "Email is not received"
End If

Set FilesObj = Nothing
Set FSO = Nothing

''Run sql to check for dupes
Set Wshell = CreateObject("wscript.shell")
Wshell.Run "cmd /c CD /d C:\ & sg2.bat", 0, false
Set Wshell = nothing

''Compare dupes found in database with baseline file of expected dupes 
Dim DupesInDB
Dim BaselineSG2
Dim FileName

FileName = "C:\logfile.txt"
DupesInDB = Readfile("C:\logfile.txt")
BaselineSG2 = Readfile("C:\baseline_sg2.txt")

'Print "DupesInDB=" & DupesinDB
'Print "BaselineSG2=" & BaselineSG2
 
Set Wshell = CreateObject("wscript.shell")
Wshell.Run "cmd /c CD /d C:\ & EmailCheckReuse.bat", 1, false
Set Wshell = nothing
