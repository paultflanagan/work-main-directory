
' HEADER
'------------------------------------------------------------------'
'    Description:  Terminates Guardian Configuration Manager, if running
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
'  - uses Repositories: GuardianConfigMgr.Main.tsr
' Post-condition:
'  - Configuration Manager is not running

' SCRIPT START

Option Explicit

reporter.ReportNote "ACTION started - Exit application"

' close any existing copy of the application
If SwfWindow("GuardianConfig_Main").Exist(0) then
	Print Now & " - Stopping Guardian..."
	SwfWindow("GuardianConfig_Main").Close
	wait 3

	' report status
	If Not SwfWindow("GuardianConfig_Main").Exist(0)  Then
		Print Now & " - Guardian stopped"
		reporter.ReportEvent micPass, "Login", "Login completed OK"
	End If
Else
	Print Now & " - No Guardian app found to be running"
End if

reporter.ReportEvent micDone, "Exit application", "ACTION completed"

' SCRIPT END
