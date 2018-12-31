'------------------------------------------------------------------
'   Description   	  :      Loads the PIM Framework spreadsheet to be used for this test
'   Project           :      Uniseries Duplicate Check 7.40.011
'   Author            :      Paul F.
'   © 2018   Systech International.  All rights reserved
'------------------------------------------------------------------
'
'	Prologue:
'   - Product Data has been loaded through FT_DualFormat_Driver via Guardian
'   
'   Epilogue:
'   - Loads the PIM spreadsheet that will run the line.

'msgbox("Artificial Breakpoint.")
call PIM.InitializeFromData()

'Preparing baseline_target.txt
'Target = baseline_740011.txt
Dim FSO, TargetFile, sTargetContents
Set FSO = CreateObject("Scripting.FileSystemObject")
Set TargetFile = FSO.OpenTextFile("C:\baseline_740011.txt", 1)			'Open file for reading
sTargetContents = TargetFile.ReadAll
TargetFile.Close

Set TargetFile = FSO.OpenTextFile("C:\baseline_target.txt", 2, True)	'Open file for writing, create file if it does not exist
TargetFile.Write(sTargetContents)
TargetFile.Close

