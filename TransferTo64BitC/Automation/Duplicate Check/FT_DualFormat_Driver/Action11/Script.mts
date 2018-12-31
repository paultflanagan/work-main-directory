' HEADER
'------------------------------------------------------------------'
'    Description:  Functional Test - Set key values for SGTIN import
'
'        Project:  Duplicate Check
'   Date Created:  2017 May
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
'   20170515   8.3.0     RichN            Initial Release
'
' 
' Pre-condition:
'  - Local dataSheet has steps defined
'  - requires Libraries: Common.Library.vbs, DotNet.Library.vbs, Guardian.Library.vbs
' Post-condition:


' START SCRIPT

Option Explicit
Print Now & " - ACTION START: Setup_SGTIN"
reporter.ReportNote "ACTION started - Duplicate Functional Test - Dual Format import;Setup SGTIN"

Dim dtStartTime : dtStartTime = Now()
Dim objWindow
Set objWindow = SwfWindow("GuardianConfig_Products")

If Not IsScreenDisplayed(objWindow.SwfLabel("lblTitleText"), "Products") Then
	NavigateToMenu objWindow.SwfTreeView("MenuTree"), "Lines and Products;Products", objWindow.SwfLabel("lblTitleText"), "Products"
End If


If FindRowInDataGrid(objWindow.SwfTable("dtgrdProduct"), "Product Name", "Product FT-D", False)  > -1 Then	' exists
	SelectRowInDataGrid objWindow.SwfTable("dtgrdProduct"), "Product Name", "Product FT-D", False				' select for update
	
	objWindow.SwfButton("btnPackagingLevels").Click
	Dim dlgLevels
	Set dlgLevels = objWindow.SwfWindow("dlgPackLevels")
	
	dlgLevels.SwfButton("btnProvisioning").Click
	Dim dlgConfig
	Set dlgConfig = dlgLevels.SwfWindow("dlgProvisioningConfig")

	Dim strValue 
	SelectRowInDataGrid dlgConfig.SwfTable("dtgConfigs"), "Format Name", "AI(01)+AI(21)", False
	strValue = dlgConfig.SwfEdit("txtKeyValue").GetROProperty("Text")
	If Left(LCase(strValue), 1) <> "z" Then ' unhidden
		dlgConfig.SwfEdit("txtKeyValue").Set "z" & strValue						' hide key value (add prefix)
		dlgConfig.SwfButton("btnSave").Click		
	End If

	SelectRowInDataGrid dlgConfig.SwfTable("dtgConfigs"), "Format Name", "SGTIN-96", False
	strValue = dlgConfig.SwfEdit("txtKeyValue").GetROProperty("Text")
	If Left(LCase(strValue), 1) = "x" Then ' hidden
		dlgConfig.SwfEdit("txtKeyValue").Set Right(strValue, Len(strValue)-1)	' unhide key value (strip prefix)
		dlgConfig.SwfButton("btnSave").Click		
	End If

	LogResult Environment("Results_File"), True, dtStartTime, Now(), "FT_DualFormat_Driver", Null, "Change Key Values for SGTIN import", "Key Values changed"

	dlgConfig.SwfButton("btnClose").Click	
	Set dlgConfig = Nothing
	dlgLevels.SwfButton("btnClose").Click	
	Set dlgLevels = Nothing
End If

objWindow.SwfButton("btnClose").Click	

reporter.ReportEvent micDone, "Duplicate Functional Test - Dual Format import;Setup SGTIN", "ACTION completed"
Print Now & " - ACTION END: Setup_SGTIN"

' END SCRIPT

