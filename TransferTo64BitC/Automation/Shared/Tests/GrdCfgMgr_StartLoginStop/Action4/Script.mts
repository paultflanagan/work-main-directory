
' HEADER
'------------------------------------------------------------------'
'    Description:  Log out of Guardian Configuration Manager
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
'  - Configuration Manager is running; user logged in
'  - uses Repositories: GuardianConfigMgr.Main.tsr
' Post-condition:
'  - Configuration Manager is running; no user logged in; login screen displayed

' SCRIPT START

Option Explicit

reporter.ReportNote "ACTION started - Logout application"

' log out of application
SwfWindow("GuardianConfig_Main").SwfButton("btnLogOut").Click

reporter.ReportEvent micDone, "Logout application", "ACTION completed"

' SCRIPT END