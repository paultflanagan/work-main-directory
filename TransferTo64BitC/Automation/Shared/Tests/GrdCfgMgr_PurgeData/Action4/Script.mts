
' HEADER
'------------------------------------------------------------------'
'    Description:  Data Purge without SPT Numbers
'
'        Project:  Guardian Configuration Manager
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
'  - Configuration Manager is running; user logged in
'  - requires Environment variables if optional Input parameters not provided: DbServer_SysadminId, DbServer_SysadminPwd
'  - requires Libraries: Common.Library.vbs, Guardian.Library.vbs
'  - uses Repositories: GuardianConfigMgr.PurgeData.tsr
' Post-condition:
'  - Configuration Manager is running; user logged in; blank screen displayed

' SCRIPT START

Option Explicit

reporter.ReportNote "ACTION started - Data Purge without SPT Numbers"

' get optional Input parameters
Dim strUser : strUser = Parameter("Input_UserAcct")
Dim strPwd : strPwd = Parameter("Input_Password")

Print Now & " - Input Parameters: UserAcct=" & strUser & "  Password=" & strPwd

' set missing parameters
If StringIsNullOrEmpty(strUser) Then
	strUser = Environment("DbServer_SysadminId")
End If
If StringIsNullOrEmpty(strPwd) Then
	strPwd = Environment("DbServer_SysadminPwd")
End If

Print Now & " - UserAcct=" & strUser & "  Password=" & strPwd

If Not IsScreenDisplayed(SwfWindow("GuardianConfig_PurgeData").SwfLabel("lblTitleText"), "Purge Production Data") Then
	NavigateToMenu SwfWindow("GuardianConfig_PurgeData").SwfTreeView("MenuTree"), "Administration;Purge Production Data", SwfWindow("GuardianConfig_PurgeData").SwfLabel("lblTitleText"), "Purge Production Data"
End If

If SwfWindow("GuardianConfig_PurgeData").SwfButton("btnNext").Exist Then	' page 1
	SwfWindow("GuardianConfig_PurgeData").SwfButton("btnNext").Click
	
	' set credentials
	SwfWindow("GuardianConfig_PurgeData").SwfComboBox("cmbAuthentication").Select "SQL Server Authentication"
	SwfWindow("GuardianConfig_PurgeData").SwfEdit("txtUser").Set strUser
	SwfWindow("GuardianConfig_PurgeData").SwfEdit("txtPwd").SetSecure strPwd
	SwfWindow("GuardianConfig_PurgeData").SwfCheckBox("chkRemoveSPT").Set "OFF"
	SwfWindow("GuardianConfig_PurgeData").SwfButton("btnNext").Click

	' confirmation
	If SwfWindow("GuardianConfig_PurgeData").Dialog("dlgConfirmPurge").Exist(5) Then	' are you sure?
		SwfWindow("GuardianConfig_PurgeData").Dialog("dlgConfirmPurge").WinButton("btnYes").Click
		wait(10)
		If SwfWindow("GuardianConfig_PurgeData").Dialog("dlgPurged").Exist(100) Then
			SwfWindow("GuardianConfig_PurgeData").Dialog("dlgPurged").WinButton("btnOK").Click	' done
			reporter.ReportEvent micPass, "Purge Data", "Purge completed"
		Else
			reporter.ReportEvent micFail, "Purge Data", "Purge may have failed"
		End If
	Else
		reporter.ReportEvent micFail, "Purge Data", "Failed to continue due to account validation"
	End If
End If ' pg 1

reporter.ReportEvent micDone, "Data Purge without SPT Numbers", "ACTION completed"

' SCRIPT END