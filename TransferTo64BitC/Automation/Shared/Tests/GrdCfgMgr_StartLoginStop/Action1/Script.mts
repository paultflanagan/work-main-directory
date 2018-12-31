
' HEADER
'------------------------------------------------------------------'
'    Description:  Starts(restarts) the Guardian Configuration Manager (as configured in UFT Environment)
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
'  - Configuration Manager may or may not be running
'  - requires Environment variables: App_ExePath, App_ExeName
'  - requires Libraries: Common.Library.vbs
'  - uses Repositories: GuardianConfigMgr.Main.tsr, GuardianConfigMgr.Login.tsr
' Post-condition:
'  - Configuration Manager is running; no user logged in; login screen displayed

'  - Login screen displayed

' START SCRIPT

Option Explicit
reporter.ReportNote "ACTION started - Start Test Application (Guardian Config Mgr)"

' close any existing copy of the application
if SwfWindow("GuardianConfig_Main").Exist(0) then
	Print "Stopping Guardian..."
	SwfWindow("GuardianConfig_Main").Close
	wait 3
End if

' get the application path+exe
Dim strAppPath : strAppPath = BuildPath(Environment("App_ExePath"), Environment("App_ExeName"))

' start the application
Print "Starting Guardian - " & strAppPath
systemutil.Run strAppPath  

' report status
If SwfWindow("GuardianConfig_Login").Exist(10) then
	reporter.ReportEvent micPass, "Start Test Application", "Test Application started OK"
	Print "Guardian started"
Else
	reporter.ReportEvent micFail, "Start Test Application", "Test Application FAILED to start"
	msgbox "Failed to start the test application"
End if

If SwfWindow("GuardianConfig_Login").SwfButton("btnLogin").Exist(0) Then
	reporter.ReportEvent micPass, "Start Test Application", "Login screen is visible"
Else
	reporter.ReportEvent micFail, "Start Test Application", "Login screen is not visible"
End If

reporter.ReportEvent micDone, "Start Test Application", "ACTION completed"

' END SCRIPT