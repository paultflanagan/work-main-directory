' HEADER
'------------------------------------------------------------------'
'    Description:  Functional Test - Enable/disable safeguards for duplication check
'
'        Project:  Duplicate Check
'   Date Created:  2017 March
'         Author:  Rich Niedzwiecki
'
'  Systech International Confidential
'  © Copyright Systech International 2017
'  The source code for this program is not published or otherwise divested of its trade secrets, 
'  irrespective of what has been deposited with the U.S. Copyright Office.
'
'      Revision History
'   Date       Version   Who              Comments         
'------------------------------------------------------------------'
'
'   20170301   8.3.0     RichN            Initial Release
'
' 
' Pre-condition:
'  - Local dataSheet has steps defined
'  - requires Libraries: Common.Library.vbs, DotNet.Library.vbs, Guardian.Library.vbs
' Post-condition:


' START SCRIPT

Option Explicit
Print Now & " PROCESSING TEST: " & datatable.Value ("Name", dtLocalsheet)
reporter.ReportNote "PROCESSING TEST: " & datatable.Value ("Name", dtLocalsheet)
Print Now & " - ACTION START: Set_DuplicateSafeGuards"
reporter.ReportNote "ACTION started - Duplicate Functional Test - Set Safeguards"

Dim dtStartTime : dtStartTime = Now()
Dim strSafeGuard, strFlag, strSQL
Dim strWhere : strWhere = " WHERE ProductId IN (SELECT ProductId FROM [Guardian].[Products] WHERE ProductName IN ('Product FT-A', 'Product FT-B', 'Product FT-C'))"
Dim boolIsChanged : boolIsChanged = False

' SafeGuard #1  (ON/OFF)
strSafeGuard = datatable.Value ("SafeGuard1", dtLocalsheet)
If Not StringIsNullOrEmpty(strSafeGuard) Then
	If LCase(strSafeGuard) = "off" Then
		strFlag = "0"
	Else
		strFlag = "1"	' default
	End If
	
	strSQL = "UPDATE [Guardian].[ProvisioningConfiguration] SET CheckUsedSPTNumbers = " & strFlag & strWhere
	ExecuteSQL GetConnectionString, strSQL, Null, Null, NULL

	LogResult Environment("Results_File"), True, dtStartTime, Now(), "Set_DuplicateSafeGuards", Null, "Set SafeGuard1 = " & strSafeGuard, "SafeGuard1 = " & strSafeGuard
	Print Now & " SafeGuard1=" & strSafeGuard
End If

' SafeGuard #2	(ON/OFF)
dtStartTime = Now()
strSafeGuard = datatable.Value ("SafeGuard2", dtLocalsheet)
If Not StringIsNullOrEmpty(strSafeGuard) Then
	If LCase(strSafeGuard) = "off" Then
		strFlag = "0"
	Else
		strFlag = "1"	' default
	End If
	
	strSQL = "UPDATE [Guardian].[ProvisioningConfiguration] SET RemoveDuplicates = " & strFlag & strWhere
	ExecuteSQL GetConnectionString, strSQL, Null, Null, NULL

	LogResult Environment("Results_File"), True, dtStartTime, Now(), "Set_DuplicateSafeGuards", Null, "Set SafeGuard2 = " & strSafeGuard, "SafeGuard2 = " & strSafeGuard
	Print Now & " SafeGuard2=" & strSafeGuard
End If

If Not boolIsChanged Then
	reporter.ReportNote "ACTION skipped - N/A"
End If

reporter.ReportEvent micDone, "Duplicate Functional Test - Set Safeguards", "ACTION completed"
Print Now & " - ACTION END: Set_DuplicateSafeGuards"

' END SCRIPT