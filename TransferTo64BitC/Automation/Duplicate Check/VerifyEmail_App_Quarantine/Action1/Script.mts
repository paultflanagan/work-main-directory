'------------------------------------------------------------------
'   Description   	  :      Determines whether the most recent lot resulted in an error by means of checking the Duplicate Check inbox for error notifications
'								In this case, emails are expected, since the lot entered has shared product numbers with the previous lot.
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
	Wshell.Run "cmd /c CD /d C:\ & EmailCheckQuarantine.bat", 1, true
	Set TextFile = FSO.OpenTextFile("C:\EmailCheckQuarantineResults.txt")
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


'Summary: Compare dupes found in database with baseline file of expected dupes 

Dim DupesInDB
Dim BaselineSG3


Dim FSOBase, FileBase
Set FSOBase = CreateObject("Scripting.FileSystemObject") 
Set FileBase = FSOBase.OpenTextFile("C:\baseline_target.txt", 1) ' 1 = ForReading

'Note: The structure for this code has been altered. Now checks the safeguard file against a baseline_target.txt file.
'This is so that this action can be used for more than just one test.
'Ensure that baseline_target.txt is given the proper contents prior to this action in the test.

BaselineSG3 = FileBase.ReadAll
FileBase.Close
Set FileBase = Nothing
Set FSOBase = Nothing
Dim bChecksBaseline
bChecksBaseline = True

If StrComp(BaselineSG3, "This test does not call for validation against baseline data.") = 0 Then
	bChecksBaseline = False
End If



'Waits until either 3rdSafeguard.txt has been updated to match the contents of baseline_target.txt, or until 6 attempts have passed
If bChecksBaseline Then
	Dim FSOCheck, FileCheck, iCounter
	Set FSOCheck = CreateObject("Scripting.FileSystemObject") 
	iCounter = 0

	Do
		'Resetting counter for next wait
		count = 0
		Set Wshell = CreateObject("wscript.shell")
		Wshell.Run "cmd /c CD /d C:\ & sg3.bat", 1, false
		Set Wshell = nothing
		Do
			Wait(5)
			count = count + 1
		Loop While Window("C:\Windows\system32\cmd.exe").Exist(1) AND count < 5

		iCounter = iCounter + 1
		'wait(10)
		Set FileCheck = FSOCheck.OpenTextFile("C:\3rdSafeguard.txt", 1) ' 1 = ForReading
		DupesInDB = FileCheck.ReadAll
		Set FileCheck = FSOCheck.OpenTextFile("C:\3rdSafeguard_" & iCounter & ".txt", 2, True) ' 2 = ForWriting, True = Create a file
		FileCheck.Write(DupesInDB)
		FileCheck.Close
	Loop While ((StrComp(DupesInDB, BaselineSG3) <> 0) AND (iCounter < 6))

	Set FileCheck = Nothing
	Set FSOCheck = Nothing

	'If test is not intended to check for baseline validation, then it should point baseline_target.txt to the baseline_null.txt file.
	If ((StrComp(DupesInDB, BaselineSG3) = 0)) Then
		LogResult Environment("Results_File"), True, dtStartTime, Now(), "Database Validation", "UNSS-3159", "Validate dupes in database", "Dupes in database validated"
	Else
		LogResult Environment("Results_File"), False, dtStartTime, Now(), "Database Validation", "UNSS-3159", "Validate dupes in database", "Dupes in database not validated"
	End If
End If
