
' HEADER
'------------------------------------------------------------------'
'    Description:  Logs into Guardian Configuration Manager using SQL credentials, 
'                  which shall be provided as either Input parameters or UFT Environment variables
'
'        Project:  Guardian Configuration Manager
'   Date Created:  2017 February
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
'   20170201   8.3.0     RichN            Initial Release
'
' 
' Pre-condition:
'  - Configuration Manager is running; no user logged in; login screen displayed
'  - optional Input parameters: Input_ServerName, Input_UserAcct, Input_Password
'  - requires Environment variables if optional Input parameters not provided: DbServer_Name, DbServer_UserId, DbServer_Pwd
'  - requires Libraries: Common.Library.vbs, Guardian.Library.vbs
'  - uses Repositories: GuardianConfigMgr.Main.tsr, GuardianConfigMgr.Login.tsr
' Post-condition:
'  - Configuration Manager is running; user logged in; blank screen displayed with menu

' SCRIPT START

Option Explicit
reporter.ReportNote "ACTION started - Login with SQL"

' verify app is running and login screen is available
If SwfWindow("GuardianConfig_Login").Exist(10) And SwfWindow("GuardianConfig_Login").SwfButton("btnLogin").Exist(0) Then
	reporter.ReportEvent micPass, "Login Test Application", "Login screen is visible"
Else
	reporter.ReportEvent micFail, "Login Test Application", "Login screen is not visible"
	ExitTest
End If

' get optional Input parameters
Dim strServer : strServer = Parameter("Input_ServerName")
Dim strUser : strUser = Parameter("Input_UserAcct")
Dim strPwd : strPwd = Parameter("Input_Password")

Print Now & " - Input Parameters: Server=" & strServer & "  UserAcct=" & strUser & "  Password=" & strPwd

' set missing parameters
If StringIsNullOrEmpty(strServer) Then
	strServer = Environment("DbServer_Name")
End If
If StringIsNullOrEmpty(strUser) Then
	strUser = Environment("DbServer_UserId")
End If
If StringIsNullOrEmpty(strPwd) Then
	strPwd = Environment("DbServer_Pwd")
End If

Print Now & " - Server=" & strServer & "  UserAcct=" & strUser & "  Password=" & strPwd

' attempt to login 
DoLogin SwfWindow("GuardianConfig_Login"), True, strUser, strPwd, strServer

' handle special case for gap file error
'If SwfWindow(testWindow).Dialog("Set Server Name").Exist(2)  Then
'	SwfWindow(testWindow).Dialog("Set Server Name").WinButton("OK").Click
'End If

' report status
If SwfWindow("GuardianConfig_Main").Exist(30) And SwfWindow("GuardianConfig_Main").SwfButton("btnLogOut").Exist(10)  Then
	reporter.ReportEvent micPass, "Login", "Login completed OK"
else
	reporter.ReportEvent micFail, "Login", "Login FAILED"
End If

reporter.ReportEvent micDone, "Login with SQL", "ACTION completed"

' SCRIPT END