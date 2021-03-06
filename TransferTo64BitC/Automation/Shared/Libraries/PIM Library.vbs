'------------------------------------------------------------------
'   Description: Functions and Subroutines PIM Lab Library         
'		                                  						                                                        
'   Project:  Master PIM Library                    
'   Date Created:  2015 March 17                                   
'         Author:  Stephen Lisa                                    
'  © 2015-2016 Systech International.  All rights reserved.             
'                                                                  
'   Revision History                                            
'   Who         			Date    			CodeVersion - Comments    
'	Stephen Lisa			Jan 01, 2015		Original File
'	Stephen Lisa			Mar 17, 2015		Updated for PBL/CBL Support
' 	Stephen Lisa			Mar 17, 2015		Added File Version Info
'	Stephen Lisa			Mar 18, 2015		Added Delay() and Delayms()
'  	Stephen Lisa			Mar 21, 2015		Added !!!! Priority Message Box
'	Stephen Lisa			Mar 24, 2015		Added {MyIP} code for SetIpsProp()
' 	Stephen Lisa			Mar 25, 2015		Added Yes/No/Cancel to user message box
'	Stephen Lisa			Mat 28, 2015		Organized RunGridRow execution Order & MsgBoxOnError
' 	Stephen Lisa			Mar 31, 2015		Removed leading, trailing and only "spaces" from cell text
' 	Lee Clark				Feb 08, 2016		PIM tests on the target with eVikIO in simulation mode
' 	Rich Niedzwiecki		Feb 09, 2016		Added DevMgrStartLot(), DevMgrStopLot(), DevMgrSuspendLot(), DevMgrResumeLot() 
'   Corey Wu                Jun 16, 2016        Added Device Recovery procedure to PostTestMsg()
'	Corey Wu				Jun 16, 2016 		Added function Minapp()that minimizes UFT when running a test
'	Corey Wu				Nov 17,	2016		Added retries to IpsLotStart() and IpsLotEnd()
'------------------------------------------------------------------
'
Const c_sPimLibraryRev = "February 9, 2016"

Dim g_oVisionSim
Dim g_oVikIo
Dim g_oPrinterSim
Dim g_bMsgBoxOnError
Dim g_bVikIoSimMode

Dim PIM
Set PIM = new cPim

Dim Vik
Set Vik = new cVik

Public Sub MyReportEvent(nEventStatus, sStepName, sDetails)

	call Reporter.ReportEvent (nEventStatus, sStepName, sDetails)
	
	If (g_bMsgBoxOnError = true) and (nEventStatus = micFail ) Then
		if msgbox(sDetails, vbOKCancel, sStepName) = vbCancel then
			Exittest
		End if
	End If
	
End Sub

Class cVik
  Private m_oVik
  Private m_nInputMask 
  Private m_nVikIndex
    
  Private Sub Class_Initialize()  
    m_nVikIndex = -1
  End Sub 

  Private Sub Class_Terminate()   
  End Sub 

  Public Function Initialize()
  	
  	g_bMsgBoxOnError = false
  	
  	'begin log script file version numbers
    MyReportEvent micPass, "PIM Library Script", c_sPimLibraryRev
    logGuardianScriptRev
    LogDeviceSimScriptRev
    ' end log script file version numbers
    
    If Not IsObject(g_oVikIo) Then
      Set g_oVikIo = CreateObject("TIPS.eVikIoApp.1")
            if Not IsObject(g_oVikIo) then
                MyReportEvent micFail, "Create eVikIo", "Failed to create eVikIo application object"
            Else
				VikIndex = 0
				g_bVikIoSimMode = 0
				on error resume next 
				g_bVikIoSimMode = g_oVikIo.VikIoSimMode
			End If
        End if 
   End Function
   
  ' Converts from 'xxxx 0010' to data and mask.
  Public Sub Pattern2Io(ByRef nRefData, ByRef nRefMask, sPattern)
    nData = 0
    nMask = 0
    For i=1 to len(sPattern)
        select case mid(sPattern,i,1)
        case "0"
            nData = nData * 2
            nMask = nMask * 2 + 1
        case "1"
            nData = nData * 2 + 1
            nMask = nMask * 2 + 1
        Case "x", "X"
            nData = nData * 2
            nMask = nMask * 2
        end select
     Next
     nRefData = nData
     nRefMask = nMask
  End Sub

  Function Io2Pattern(ByVal nData, ByVal nMask)
    sPattern = ""
    nDigits = 0
    do while nMask <> 0
    	If (nMask And 1)  Then
           if (nData And 1) = 1 then
               sPattern = "1" & sPattern
           else
               sPattern = "0" & sPattern
           end if
    	Else
    		sPattern = "x" & sPattern
    	End If
    	
        nMask = nMask \ 2
        nData = nData \ 2
        
        if (nDigits And 3) = 3 and nMask <> 0 then
            sPattern = " " + sPattern
        end if
        nDigits = nDigits + 1
     loop
     
     Do while (nDigits And 3) <> 0 
   		sPattern = "x" & sPattern
        nDigits = nDigits + 1
     loop
     
     Io2Pattern = sPattern
  End Function  
  
    Public Property Get ResultLineCount
        ResultLineCount = m_oVik.All("ResultLineCount").TextValue
        
    End Property

    ' example   Vik.ResultLineCount = "2 2 0 0"
    Public Property Let ResultLineCount(newVal)
        m_oVik.All("ResultLineCount").TextValue = newVal
    End Property

    Public Property Get VikIndex
        VikIndex = m_nVikIndex
    End Property

    Public Property Let VikIndex(newVal)
        newVal = CInt(newVal)
        If m_nVikIndex <> newVal Then
            m_nVikIndex = newVal
            Set m_oVik = g_oVikIo.Viks(nVik)
            if Not IsObject(m_oVik) then
                MyReportEvent micFail, "Create VIK", "Failed to create Vik object"
            end if
        
            m_oVik.All("UpdateMode").TextValue = "UpdateModeInputs"
            m_oVik.OutputForceData = 0
            m_oVik.OutputForceMask = 0
            m_nInputMask = &H0fff		
        End If
    End Property

    ' get the OutputState of the 'system under test'
	Public Property Get VikInputState
		If g_bVikIoSimMode <> 0 Then
			' simulation mode: directly read the outputs
			VikInputState = m_oVik.OutputState
		Else
			' real io: real outputs of the 'system under test' are wired to 
			' real inputs of the system performing the tests.
			VikInputState = m_oVik.InputState
		End If
	End Property
      
    Public Property Get InputState
        InputState = Io2Pattern(m_oVik.InputState, m_nInputMask)
    End Property
    
    Public Property Get InputMask
        InputMask = Io2Pattern(m_nInputMask, m_nInputMask)
    End Property

    Public Property Let InputMask(newVal)
        Dim nData, nMask
        call Pattern2Io(nData, nMask, newVal)
        m_nInputMask = nMask
    End Property
	
    ' set the InputState of the 'system under test'
    Public Property Get OutputForceMask
    	Dim nMask
		If g_bVikIoSimMode <> 0 Then
			' simulation mode: running on the "system under test", 
			' read the inputs directly
			nMask = m_oVik.InputForceMask
		Else
			' real io: real inputa of the 'system under test' are wired to 
			' real outputa of the system performing the tests.
			' allow the real outputs to be used.
			nMask = m_oVik.OutputForceMask
        End if
        OutputForceMask = Io2Pattern(nMask, nMask)
    End Property

    ' set the InputState of the 'system under test'
    Public Property Let OutputForceMask(newVal)
        Dim nData, nMask
        call Pattern2Io(nData, nMask, newVal)
        
		If g_bVikIoSimMode <> 0 Then
			' simulation mode: running on the "system under test", 
			' read the inputs directly
			m_oVik.InputForceMask = nMask
		Else			
			' real io: real input of the 'system under test' are wired to 
			' real output of the system performing the tests.
			' allow the real outputs to be used.
	        m_oVik.OutputForceMask = nMask
	    End if 
    End Property
    
    ' set the InputState of the 'system under test'
    Public Property Get OutputForceData
		If g_bVikIoSimMode <> 0 Then
			' simulation mode: running on the "system under test", 
			' read the inputs directly
			OutputForceData = Io2Pattern(m_oVik.InputForceData, m_oVik.InputForceMask)
		Else
			' real io: real input of the 'system under test' are wired to 
			' real output of the system performing the tests.
			' allow the real outputs to be used.
			OutputForceData = Io2Pattern(m_oVik.OutputForceData, m_oVik.OutputForceMask)
		End if 
    End Property
	
    ' set the InputState of the 'system under test'
    Public Sub SetOutputForceData(sStepName,newVal)
    
        Dim nData, nMask
        call Pattern2Io(nData, nMask, newVal)
        
		If g_bVikIoSimMode <> 0 Then
			' simulation mode: running on the "system under test", 
			' force the inputs directly
	        If (nMask <> m_oVik.InputForceMask) or (nData <> m_oVik.InputForceData) Then
				m_oVik.InputForceMask = nMask
				m_oVik.InputForceData = nData
				sMsg = " - Output Set: " & Io2Pattern(nData, nMask)
				call MyReportEvent(micPass, sStepName, sMsg)
	        End If
		Else
			' real io: real inputs of the 'system under test' are wired to 
			' real outputs of the system performing the tests.
			' allow the real outputs to be used.
	        If (nMask <> m_oVik.OutputForceMask) or (nData <> m_oVik.OutputForceData) Then
				m_oVik.OutputForceMask = nMask
				m_oVik.OutputForceData = nData
				sMsg = " - Output Set: " & Io2Pattern(nData, nMask)
				call MyReportEvent(micPass, sStepName, sMsg)
	        End If
		End if

	End Sub
       
    Public Sub VerifyInput(sExpectedPattern, sStepName)
        Dim nData, nMask, nMaxTrys, bAbortOnError
        Dim sBits1316, sOriginalExpectedPattern
        
        sBits1316 = ""
                
        nMaxTrys = 200
        bAbortOnError = false 
        
        ' override default number to trys to increase timeout
        ' 10ms per loop when polling for IO results.
		If Left(sExpectedPattern,Len("Wait(")) = "Wait(" Then
			aTempStrings = split(sExpectedPattern, ")")
			sExpectedPattern = aTempStrings(1)
			nMaxTrys = Int(Mid(aTempStrings(0),6))
			nMaxTrys = nMaxTrys /10
			bAbortOnError = true & m_bAbortOnError
		End If
		
		sOriginalExpectedPattern = sExpectedPattern
		
		sExpectedPattern = replace(sExpectedPattern, " ", "")   
		If len(sExpectedPattern) > 15 and g_bVikIoSimMode = 0 Then
             	
        	' there are only 12 inputs mapped to 16 outputs
        	sBits1316 = Left(sExpectedPattern, 4)
        	If InStr (sBits1316, "1") > 0 or  InStr(sBits1316, "0") > 0  Then
        		' if 13-16 is 1 or 0 then swap with bits 1-4
        		sExpectedPattern = "XXXX" &  Left(Right(sExpectedPattern, 12),8) & sBits1316
        	else
        		sBits1316 = ""
        	End If
        	
        End if
		       
        call Pattern2Io(nData, nMask, sExpectedPattern)
        
        nOldValue = VikInputState And m_nInputMask
        
        sTempPattern = Io2Pattern(nOldValue, nMask)
        If sBits1316 <> "" Then
         	' flip bits 13-16 and 1-4
        	sTempPattern = Right(sTempPattern, 4) & " " & Left (sTempPattern, 9) & " XXXX"
         End If
                
        sFail = " Received: " & sTempPattern
        nPassCode = micFail
        
        For nTry = 1 To nMaxTrys
            nNewValue = VikInputState And nMask
            If nData = nNewValue Then
                nPassCode = micPass
                ' wait for DV to drop 
				'For nWait = 1 To 5 Step 1
				'	If nData <> VikInputState And nMask Then
				'		Exit For
				'	End If
				'	Wait 0, 100
				'Next
                Exit For
            ElseIf nOldValue <> nNewValue Then
                nOldValue = nNewValue
                sTempPattern = Io2Pattern(nOldValue, nMask)
                If sBits1316 <> "" Then
                 ' flip bits 13-16 and 1-4
                 sTempPattern = Right(sTempPattern, 4) & " " & Left (sTempPattern, 9) & " XXXX"
                End If
                sFail = sFail & ",  " & sTempPattern				
            End If
            Wait 0, 10
        Next
        
        sMsg = " - Expected: " & sOriginalExpectedPattern
        If nPassCode = micFail Then
            sMsg = sMsg & sFail 
        else
        	bAbortOnError = false
        End If
        
        call MyReportEvent (nPassCode, sStepName, sMsg)
        
        If bAbortOnError = true Then
        	ExitTest 
        End If
    End Sub
    
 End Class



Class cPim
	Private m_nDataValidIndex
	Private m_sIpsPimFiles
	Private m_sTestName
	Private m_bAbortEnabled
	Private m_bNoMsgBox

	Private Sub Class_Initialize() 
		DataValidIndex = 0
		VisionSim   	
	End Sub  

    Private Sub Class_Terminate()   
    End Sub 
    
    Sub Initialize()
    End Sub
    
    Public Property Get VisionSim
   		If Not IsObject(g_oVisionSim) Then
	    	Set g_oVisionSim = CreateObject("VisionSimulator.Application")
			if Not IsObject(g_oVisionSim) then
				MyReportEvent micFail, "Create VisionSimulator", "Failed to create VisionSimulation application object"
				ExitTest
			end if
		End If
		Set VisionSim = g_oVisionSim
    End Property
    
    Public Property Get PrinterSim
   		If Not IsObject(g_oPrinterSim) Then
	    	Set g_oPrinterSim = CreateObject("DeviceSimulator.CSharpServerObject")
			if Not IsObject(g_oPrinterSim) then
				MyReportEvent micFail, "Create Printer Simulator", "Failed to create Printer Simulation application object"
				ExitTest
			end if
		End If
		Set PrinterSim = g_oPrinterSim
    End Property

	Sub InitializeFromData()
	
		call PIM.Initialize()
		PIM.DataValidIndex = 0

		Call Vik.Initialize()
		Vik.VikIndex = 0
		Vik.InputMask = "1111 1111 1111"
		
		
		Call DataTable.SetCurrentRow(nRow)
			
		sRemoteHost = DataTable.Value("RemoteHost", DtLocalSheet)
		sIpsPimFiles = DataTable.Value("IpsPimFiles", DtLocalSheet)
		sTestDataFile = DataTable.Value("TestDataFile", DtLocalSheet)
		sTestDataSheets = DataTable.Value("TestDataSheets", DtLocalSheet)
		sGuardianSerialNumbers = DataTable.Value("GuardianSerialNumbers", DtLocalSheet)
		sVisionProject = DataTable.Value("VisionProject", DtLocalSheet)
		sTestOverrides = DataTable.Value("TestOverrides", DtLocalSheet)
		
		If Instr(sTestOverrides, "AbortOnFail") > 0 Then
			' abort test if Wait() command failed
			m_bAbortOnError = true
		Else	
			m_bAbortOnError = false
		End If
		
		If Instr(sTestOverrides, "NoMsgBox") > 0 Then
			' hide message boxes
			m_bNoMsgBox= true
		Else
			m_bNoMsgBox = false
		End If
		
		If Instr(sTestOverrides, "MsgBoxOnError") > 0 Then
			' display a message box on any failed test
			g_bMsgBoxOnError= true
		Else
			g_bMsgBoxOnError = false
		End If
		
		If len(sTestDataSheets) = 0  Then
			sTestDataSheets = "Setup,Good1,CleanUp,GridData"
		End If
		
		' purge Guardian and reload serial numbers		
		If len(sGuardianSerialNumbers) > 0 Then
			'load new serial numbers
			Call InitializeGuardian("C:\PIMLabTestData\GuardianFiles\", sGuardianSerialNumbers)
		End If
		
		' Connect to IPS Test Station 
		if len(sRemoteHost) > 0 then
			PIM.RemoteHost = sRemoteHost
		End if
		
		' load IPS Files into IPS Engine
		If len(sIpsPimFiles) > 0 Then
			PIM.IpsPimFiles = sIpsPimFiles
		End If
		
		' load vision project
		If len(sVisionProject) > 0 Then
			sVisionProject = "D:\tips\pimlab\" & sVisionProject
			Call VisionOpen(sVisionProject)
		End If
		
		If Len(sTestDataFile) > 0 Then
			If Instr(sTestDataFile, "\") = 0 Then
				' use PIM Lab Default Path
				sTestDataFile = "C:\PIMLabTestData\" & sTestDataFile
			End If
		
			'nYesNo = Msgbox ("Load PIM Test data sheet from file?" & vbCRLF & "'" & sTestDataFile & "'", vbYesNoCancel, "Load Test Data")	
			nYesNo = vbYes			
			If  nYesNo = vbYes Then
				aSheetNames = split(sTestDataSheets, ",")
				For nSheetIndex = 0 To uBound(aSheetNames)
				    If aSheetNames(nSheetIndex) = "GridData" Then
				    	call datatable.ImportSheet(sTestDataFile, aSheetNames(nSheetIndex), "Global")
				    else
						call datatable.ImportSheet(sTestDataFile, aSheetNames(nSheetIndex), aSheetNames(nSheetIndex))
					End if 
					MyReportEvent micPass, "Data Sheet Load", sTestDataFile + ":" + aSheetNames(nSheetIndex)
				Next
				
			End If
			
'			If nYesNo = vbCancel Then
'				call MyReportEvent( micFail, sStepName, "Test ABORTED by operator")
'				ExitTest  ' abort the test 
'			End If
     
            If len(IpsPimFiles) > 0 Then
				 Call PIM.IpsOpen("D:\\tips\\pimlab\\", IpsPimFiles)	
			End If
			
		End If	
		
	End Sub
	
	Public Property Let IpsPimFiles(sFiles)
		sFiles = replace (sFiles, vbCR, "") 
		sFiles = replace (sFiles, vbLF, ",")
		m_sIpsPimFiles = sFiles
	End Property

	Public Property Get IpsPimFiles
		IpsPimFiles = m_sIpsPimFiles
	End Property
	
	Public Property Let RemoteHost(newVal)
		VisionSim.RemoteHost = newVal
	End Property

  	Public Property Get RemoteHost
		RemoteHost = VisionSim.RemoteHost
	End Property
	
	Public Property Let DataValidIndex(newVal)
		m_nDataValidIndex = CLng(newVal)
	End Property

  	Public Property Get DataValidIndex
		DataValidIndex = m_nDataValidIndex
	End Property
	
	Public Property Let CommSettings(newVal)
	' example setting of serial port settings
	' PIM.CommSettings = "com1: baud=19200 parity=n data=8 stop=1"
        VisionSim.CommSettings = newVal
    End Property

    Public Property Get CommSettings
        CommSettings = VisionSim.CommSettings
    End Property
     
    Public Sub CommWrite(sData)
    ' example sending of serial data, can use all the control code names (ascii chart), decimal, or hex in the strings
	' call PIM.CommWrite( "{stx}1234567{\xd}{etx}" )

        sMsg = VisionSim.CommWrite(sData)
        If len(sMsg)>0 And Left(sMsg,1) <> "*" Then
            MyReportEvent micPass, "CommWrite Success", sMsg
        Else
            MyReportEvent micFail, "CommWrite Failed", sMsg
            ExitTest
        End If
    End Sub     
     
    Public Sub CommWriteRemote(sData)
    ' example sending of serial data, can use all the control code names (ascii chart), decimal, or hex in the strings
	' call PIM.CommWriteRemote( "{stx}0114011221ABC{gs}AAAA{12}{etx}" )
	
        sMsg = VisionSim.CommWriteRemote(sData)
        If len(sMsg)>0 And Left(sMsg,1) <> "*" Then
            MyReportEvent micPass, "CommWriteRemote Success", sMsg
        Else
            MyReportEvent micFail, "CommWriteRemote Failed", sMsg
            ExitTest
        End If
    End Sub
	
	Public Sub PostVisionResults(sResultPattern, sData)
        Dim nData, nMask
        call Vik.Pattern2Io(nData, nMask, sResultPattern)
		call VisionSim.PostVisionResults(m_nDataValidIndex, nData, nMask, sData)
	End Sub
	
	Public Sub PostGridVisionResults(nStationIndex, sGridData)
     	' Run <rows>,<coLs>,<read codes>,<r1>,<c1>,<bc1>,<r2><c2><bc2>,....
	    ' ex:  Run 2,3,2,1,1,epc1,1,2,ecp2
        call VisionSim.PostPackByLayer(nStationIndex, sGridData)
	End Sub
	
	Public Sub PostImageVisionResults(sStepName, sImageFiles)
     	' Load image files to test - 
	    call VisionTestImageSet(sStepName, "S1", "D:\tips\pimlab\Images\", sImageFiles, 250, "", "")
	End Sub
		
	Public Sub VisionOpen(sProjectPath)
        sMsg = VisionSim.VisionOpen(sProjectPath)
        If len(sMsg)>0 And Left(sMsg,1) <> "*" Then
        	MyReportEvent micPass, "Vision Open", sProjectPath
        Else
            MyReportEvent micFail, "Vision Open Failed", sMsg
            ExitTest
        End If
    End Sub
                
    Public Sub VisionTestImageSet(sStepName, sStationName, sImageDirectory, sImageSelect, nImageDelay, sLogPropertyNames, sLogFilename)
        sMsg = VisionSim.VisionTestImageSet(sStationName, sImageDirectory, sImageSelect, nImageDelay, sLogPropertyNames, sLogFilename)
        If len(sMsg)>0 And Left(sMsg,1) <> "*" Then
        	MyReportEvent micPass, sStepName, sImageDirectory
        Else
            MyReportEvent micFail, sStepName, sMsg
            ExitTest
        End If
	End Sub


	Public Sub IpsSetPropValue(sPropList)
		sPropList = replace (sPropList, vbCR, "") 
		sPropList = replace (sPropList, vbLF, "&")
		
		' get IP of the UFT test station
		sPropList = replace (sPropList, "{MyIp}", FindIp())
		' sPropList: Ampersand separated 'Name=Value' pairs
		' ex: TrigStep1=1&TrighStep2=2
		
        call VisionSim.SetIpsProperties(sPropList)
	End Sub 
	
	Public Sub IpsExecCommand(sCommand)
	
		If (Left(sCommand, Len("StartLot")) = "StartLot") or (Left(sCommand, Len("ProjectRun")) = "ProjectRun") Then
			Call PIM.IpsLotStart(200000)
			
		ElseIf Left(sCommand, Len("StopLot")) = "StopLot" Then
			If Window("IpsEngine").WinToolbar("ToolbarWindow32").Exist(2) Then
				Window("IpsEngine").WinToolbar("ToolbarWindow32").Press 6
			Else
			    For nTry = 1 To 10 Step 1
			    	Wait 1
			    	If Window("IpsEngine").WinToolbar("ToolbarWindow32").Exist(2) Then
			    		Window("IpsEngine").WinToolbar("ToolbarWindow32").Press 6
			    		Exit For
			    	End If
			    Next
			End If
			
		ElseIf Left(sCommand, Len("EndLot")) = "EndLot" Then
			Call PIM.IpsLotEnd(200000)
			
		Elseif Left(sCommand, Len("DeviceRecovery")) = "DeviceRecovery" Then
			Call PIM.IpsDeviceRecovery(200000)
			Wait 2
		ElseIf Left(sCommand, Len("ResultLineCount")) = "ResultLineCount" Then
			Dim eqCharPos
			Dim value
			eqCharPos = Instr(1, sCommand, "=")
			value = Trim(Mid(sCommand, eqCharPos+1))
			'msgbox value
			'msgbox Vik.ResultLineCount
			Vik.ResultLineCount = value
			
		ElseIf Left(sCommand, Len("CommSettings=")) = "CommSettings=" Then
			PIM.CommSettings= Mid(sCommand, Len("CommSettings=") + 1)
		
		ElseIf Left(sCommand, Len("CommWrite=")) = "CommWrite=" Then
			Call PIM.CommWrite (Mid(sCommand, Len("CommWrite=") + 1))
			
		ElseIf Left(sCommand, Len("CommWriteRemote=")) = "CommWriteRemote=" Then
			Call PIM.CommWriteRemote (Mid(sCommand, Len("CommWriteRemote=") + 1))
		
		ElseIf Left(sCommand, Len("LotControlStartLot=")) = "LotControlStartLot=" Then
			Dim aPropList, aPropPairs
			sPropList = Mid(sCommand, Len("LotControlStartLot=") + 1)
			sPropList = replace (sPropList, vbCR, "") 
			sPropList = replace (sPropList, vbLF, "&")
			aPropPairs = split(sPropList, "|")
			Call PIM.DevMgrStartLot(aPropPairs(0),aPropPairs(1))
			
		ElseIf Left(sCommand, Len("LotControlEndLot")) = "LotControlEndLot" Then
			Call PIM.DevMgrStopLot()
			
		ElseIf Left(sCommand, Len("LotControlSuspendLot")) = "LotControlSuspendLot" Then
			Call PIM.DevMgrSuspendLot()
			
		ElseIf Left(sCommand, Len("LotControlResumeLot=")) = "LotControlResumeLot=" Then
			Call PIM.DevMgrResumeLot(Mid(sCommand, Len("LotControlResumeLot=") + 1))
			
		Else
			IpsSetPropValue(sCommand)
		End If
		
	End Sub
                                
	Public Sub VerifyIpsPropValue(sStepName, sPropList)
	
		' sPropList: Ampersand or line separated 'Name=Value' pairs
		' ex: TrigStep1=1&TrighStep2=2
		
		Dim aPropPairs, aPropAndValue, nPropIndex, sName, sValue
		Dim sPropNames, sPropReturns, nPassCode, sMsg
		Dim bWaitUntil, nTryMax
		bWaitUntil = false
		nTryMax = 200
		
		' check if infinite wait requested
		If Left(sPropList, Len("WaitUntil:")) = "WaitUntil:" Then
			Print sPropList
			sPropList = Mid(sPropList, Len("WaitUntil:") + 1)
			bWaitUntil = true
		End If
		
		sPropList = replace (sPropList, vbCR, "") 
		sPropList = replace (sPropList, vbLF, "&")
		
		' get IP of the UFT test station
		sPropList = replace (sPropList, "{MyIp}", FindIp())
		
		aPropPairs = split(sPropList, "&")
		For nPropIndex = 0 To UBound(aPropPairs)
		
			aPropAndValue = split(aPropPairs(nPropIndex), "=")
			If UBound(aPropAndValue) <> 1 Then
				Call MyReportEvent(micFail, sStepName, sMsg)
				ExitTest
			End If
			
			sName = trim(aPropAndValue(0))
			sValue = trim(aPropAndValue(1))
			
			If nPropIndex = 0 Then
				sPropNames = sName
				sExpected = sName & "=" & sValue
			Else
				sPropNames = sPropNames & "," & sName
				sExpected = sExpected & "&" & sName & "=" & sValue
			End If
		Next
		
		' poll for IPS Property to check
		If bWaitUntil Then	
			'Dim sLastReturn, sLastReturnPrior
			
			Do	' check every minute indefinitely 
				sPropReturns = trim(VisionSim.GetIpsProperties(sPropNames))			
				Print sPropReturns
				If sExpected = sPropReturns Then
					Exit Do
				End If
				'If (sPropReturns = sLastReturn) And (sLastReturn = sLastReturnPrior) Then
				'	nMinutesWaiting = nMinutesWaiting + 1
				'End If
	            'sLastReturnPrior = sLastReturn 
	            'sLastReturn = sPropReturns
	            Wait 60, 0
			Loop
		Else
	        For nTry = 1 To nTryMax
				sPropReturns = trim(VisionSim.GetIpsProperties(sPropNames))
				If sExpected = sPropReturns Then
					Exit For
				End If
	            Wait 0, 10
			Next 
		End If
		
		If sExpected = sPropReturns Then
	       	nPassCode = micPass
			sMsg = "Expected: '" & sExpected & "' == " & "Received: '" & sPropReturns & "'"
		else
	       	nPassCode = micFail
			sMsg = "Expected: '" & sExpected & "' <> " & "Received: '" & sPropReturns & "'"
		End If 
		
		Call MyReportEvent(nPassCode, sStepName, sMsg)

	End Sub
	
	Public Sub RunVisionTestStep(sStepName, sVStationData, sVisionResultLines, sExpectedOutputPattern, sMachineSignalPattern)
		'initialize loop 
		Dim aVisionData
		nLoop = 1 
			
		sVStationData = Trim(sVStationData)
	
		 ' Loop test step and increament test data by +
		If Mid(sVStationData,1,5) = "Loop(" Then
			aLoopStrings = split(sVStationData, ")")
			sVStationData = aLoopStrings(1)
			nLoop = Int(Mid(aLoopStrings(0),6))
			
			sVStationData = Trim(sVStationData)
					
		End If
		
		For Iterator = 1 To nLoop Step 1
		    
		    If len(sVisionResultLines) Or len(sVStationData) Then
				'  send data to vision station simulator
		    	sParameters = sVStationData 
				call PostVisionResults(sVisionResultLines, sParameters)	
			End If
			
			If nLoop > Iterator Then
				'check outputs for each iterations
				'skip the last check -  The check will be done by the main execution order
				call CheckExpectedOutputs(sStepName, sExpectedOutputPattern, sMachineSignalPattern)
			End If
				
			' increament serial number by +1 if looping
			If len(sVStationData) > 0 and nLoop > 1 then
				aVisionData = split(sVStationData, "&")
				sVStationData = ""
			    For nPropIndex = 0 To UBound(aVisionData)
		    
			    	If nPropIndex > 0 Then
			    		'append the & delimiter before adding the next string
			    		sVStationData = sVStationData & "&"
			    	End If
			    	
					sTemp = aVisionData(nPropIndex)	
					sInc = "0000" & (Int(Right(sTemp, 4)) + 1)
					sTemp = Mid(sTemp, 1, len(sTemp) - 4) & Right(sInc, 4)
					sVStationData = sVStationData &  sTemp
				Next

			End if
		Next
	
	End Sub
	
	Public Sub RunVisionImageTestStep(sStepName, sImageFiles)
		' load one or more images to test
		Call PostImageVisionResults(sStepName, sImageFiles)
	End Sub
	
	Public Sub RunVisionGridTestStep(sStepName, sGridName)
		'initialize loop counts
		sGridName = Trim(sGridName)
			 
		sData = "Run" 
		nCellCount = 0
				
		nSheetRowCount = DataTable.GlobalSheet.GetRowCount
		
		For nSheetRow = 1 To nSheetRowCount
			' Find the Grid Data from the Global Sheet  no spaces are allowed in the GridName
			call DataTable.GlobalSheet.SetCurrentRow(nSheetRow)
			If Trim( DataTable.Value("GridName", DtGlobalSheet) )= sGridName then
				nGridRowCount = CInt(DataTable.Value("GridRows", DtGlobalSheet))
				nGridColCount = CInt(DataTable.Value("GridColumns", DtGlobalSheet))
				
				sData = sData & "," & nGridRowCount & "," & nGridColCount & ",<CellCount>"
				
				'store Row1/Col1 as referenc string 
				sStartValueHeader = Trim(DataTable.Value(4, DtGlobalSheet))
				nStartValueInt = Int(Right(sStartValueHeader, 5)) ' 5 right digits only.
				sStartValueHeader = Mid(sStartValueHeader, 1, len(sStartValueHeader) - 5) 
				
				' Build the Grid Message
			   	For nRow = 0 To nGridRowCount-1
			   		For nCol = 0 To nGridColCount-1
						sCellData = Trim(DataTable.Value(nCol+4, DtGlobalSheet))
						If Len(sCellData)>0 Then
							nCellCount = nCellCount + 1
							If Left(sCellData,1) = "+" Then
								' shorthand notation to add value to first string.
								sInc = "00000" & (nStartValueInt + Int(Mid(sCellData, 2)))
								sCellData = sStartValueHeader & Right(sInc,5)
							End If
							sData = sData & "," & nCol & "," & nRow & "," & sCellData
						End If
			   		Next
			   		DataTable.GlobalSheet.SetNextRow
			   Next
			   
			   Exit For
			End If
		Next
			   
		If nCellCount > 0 Then
		    sData = Replace(sData, "<CellCount>", nCellCount)
		    call PostGridVisionResults(m_nDataValidIndex, sData)
		else
            call MyReportEvent(micFail, sStepName, "Grid '"& sGridName & "' not found in row " & DataTable.GetCurrentRow)
			ExitTest
		End If
	
	End Sub
	
	Public Sub RunLabelTestStep(sStepName, sExpectedLabelData)
		Dim sLastLabel	
		
		' check the string from the printers emulator
		g_oPrinterSim="none"
		
		sLastLabel = PrinterSim.PrintData
		
		If sExpectedLabelData = sLastLabel Then
           	nPassCode = micPass
			sMsg = "Expected: " & sExpectedLabelData & " == " & "Received: " & sLastLabel 
		else
            nPassCode = micFail
            sMsg = "Expected: " & sExpectedLabelData & " <> " & "Received: " & sLastLabel
		End If 
		
		Call MyReportEvent(nPassCode, sStepName, sMsg)
		
	End Sub
	
	Public Sub CheckExpectedOutputs(sStepName, sExpectedOutputPattern, sMachineSignalPattern)
		
		If len(sMachineSignalPattern) > 0 Then
			' set the machine output pattern for eS-VIK Inputs
			call VIK.SetOutputForceData(sStepName, sMachineSignalPattern)
		End If
		
		If len(sExpectedOutputPattern) > 0 Then
			' check expected results from eS-VIK Outputs	
			call VIK.VerifyInput(sExpectedOutputPattern, sStepName)
		End If
		
	End Sub
	
	Public Sub RunGridRow()
		
		sTestName = trim(DataTable.Value("TestName", DtLocalSheet))
		If Len(sTestName) > 0 Then
			m_sTestName = sTestName
		End If
				
		' test step description
		sStepName = trim(DataTable.Value("StepName", DtLocalSheet))
		If Len(sStepName) = 0 Then
			' missing required field
			call Msgbox ("Missing Step Name. Test Aborted.")
			MyReportEvent micFail, "????", "Step Name Not Defined.  Test Aborted."	
			ExitTest
		End If
		
		sStepName = m_sTestName & ":" & sStepName
		
		' data for vision station DataName=nnnnnnnnnnnnnnnn
		sVStationData = trim(DataTable.Value("vData", DtLocalSheet))

		' results for vision station  (16 bits) xxxx xxxx xxxx xxxx
		sVisionResultPattern = trim(DataTable.Value("vResults", DtLocalSheet))

		' Data Valid Index
		sDataValidIndex = trim(DataTable.Value("vDV", DtLocalSheet))

		' expected VIK Output Pattern (12 bits) xxxx xxxx xxxx
		sExpectedOutputPattern = trim(DataTable.Value("ExpectedOutput", DtLocalSheet))

		' expected Label data string 0|f1|f2|f3...
		sExpectedLabelData = trim(DataTable.Value("ExpectedLabelData", DtLocalSheet))

		' Set Machine Signals eS-VIK Inputs xxxx xxxx xxxx xxxx 
		sSetInputs = trim(DataTable.Value("SetInputs", DtLocalSheet))

		' display a message before executing each test step
		sPostTestMessage = trim(DataTable.Value("PostTestMessage", DtLocalSheet))

		' Set IPS Propert Alias
		sSetIpsPropValue = trim(DataTable.Value("IpsSetPropValue", DtLocalSheet))

		' Check expected value for IPS Property after all other operations
		sExpectedIpsPropValue = trim(DataTable.Value("ExpectedIpsPropValue", DtLocalSheet))

		'  run test steps
		If len(sDataValidIndex) > 0 Then
			PIM.DataValidIndex = sDataValidIndex
		End If
		
		'set IPS Property value
		If len(sSetIpsPropValue) > 0 Then
			call PIM.IpsExecCommand(sSetIpsPropValue)
		End If

		' send data to vision station
		Call SendVisionData(sStepName, sVStationData, sVisionResultPattern, sExpectedOutputPattern, sSetInputs)
		
		' check expected outputs
		Call PIM.CheckExpectedOutputs(sStepName, sExpectedOutputPattern, sSetInputs)
	
		' check ips property value
		If len(sExpectedIpsPropValue) > 0 Then
			call PIM.VerifyIpsPropValue(sStepName, sExpectedIpsPropValue)
		End If
			
		' check label data
		Call CheckLabelData(sStepName, sExpectedLabelData)
			
		' post step message	
		call PostTestMsg(sStepName, sPostTestMessage)
	End Sub
	
	Private Sub SendVisionData(sStepName, sVStationData, sVisionResultPattern, sExpectedOutputPattern, sSetInputs)
		If len(sVStationData) > 0 Then ' send data to vision station 1

			if Instr(sVStationData, "Grid=") then
				' run a Vision Grid Data test  Grid=<GridName> 
				' data name is look up from Global Sheet
				sVStationData = replace (sVStationData, "Grid=", "") 
				Call PIM.RunVisionGridTestStep(sStepName, sVStationData)
			ElseIf Instr(sVStationData, "Image=") Then
				' run a Vision Grid Data test  Grid=<GridName> 
				' data name is look up from Global Sheet
				sVStationData = replace (sVStationData, "Image=", "") 
				Call PIM.RunVisionImageTestStep(sStepName, sVStationData)
			Else
			
				' run a Vision IPS Data test
				' replace LF with & for multi-datanames in IpsVisionResults
				sVStationData = replace (sVStationData, vbCR, "") 
				sVStationData = replace (sVStationData, vbLF, "&")
	   			call PIM.RunVisionTestStep(sStepName, sVStationData, sVisionResultPattern, sExpectedOutputPattern, sSetInputs)
   			End if
   		
   		End if
	End Sub
	
	Private Sub CheckLabelData(sStepName, sExpectedLabelData)
		if len(sExpectedLabelData) > 0 Then ' or send printer trigger and check label data
			If Left(sExpectedLabelData,2)="**" AND Right(sExpectedLabelData,2)="**" Then
				arrArgs = Split(sExpectedLabelData, "**")
				sCommand = arrArgs(1)
				If sCommand = "PRINTERSETUP" Then
					strIP = findIP()
					deviceName = arrArgs(2)
					portNum = arrArgs(3)
					autoPrint = arrArgs(4)
					DeviceSimSetup strIP, deviceName, portNum, autoPrint
				ElseIf sCommand = "PRINT" Then
					DeviceSimPrint
				ElseIf sCommand = "OPENCONNECTION" Then
					DeviceSimOpenConnection
				ElseIf sCommand = "CLOSECONNECTION" Then
					DeviceSimCloseConnection				
				ElseIf sCommand = "ENABLE" Then
					functionName = arrArgs(2)
					DeviceSimProtocolCheck functionName, "True"
				ElseIf sCommand = "DISABLE" Then
					functionName = arrArgs(2)
					DeviceSimProtocolCheck functionName, "False"
				ElseIf sCommand = "CLOSE" Then
					DeviceSimClose	
				ElseIf sCommand = "CLEAR" Then
					DeviceSimClearPrintData						
				End If					
			Else
				'replace LF with "|" for Printer Match String
				sExpectedLabelData = replace (sExpectedLabelData, vbCR, "") 
				sExpectedLabelData = replace (sExpectedLabelData, vbLF, "|")
				If Right(sExpectedLabelData, 1) <> "|" Then
					sExpectedLabelData = sExpectedLabelData & "|"
				End If
				Call PIM.RunLabelTestStep(sStepName, sExpectedLabelData)			
			End If
		End if
	End Sub
	
	Private Sub PostTestMsg(sStepName, sMessage)
		Dim nDelay, sCmd
		Dim aTempStrings
		
		nDelay = 0
		sCmd = ""
		If len(sMessage) > 0 Then
			' Delay command in seconds
			sCmd = "Delay("
			If InStr(sMessage, sCmd) > 0 Then
				sMessage = trim(sMessage)
				aTempStrings = split(sMessage, ")")
				nDelay = Int(Mid(aTempStrings(0),Len(sCmd)+1))
				wait nDelay, 0
			elseIf InStr(sMessage, "Delayms(") > 0 Then
				' Delay command in milliseconds
				sCmd = "Delayms("
				aTempStrings = split(sMessage, ")")
				nDelay = Int(Mid(aTempStrings(0),Len(sCmd)+1))
				wait 0, nDelay
			ElseIf InStr(sMessage, "DRT") > 0 Then
				If Window("IpsEngine").Dialog("IPS Engine Stop").Exist(5) Then
					prnErrMsg = Window("IpsEngine").Dialog("IPS Engine Stop").WinEditor("PrintErrorMessage").GetROProperty("text")
					'msgbox prnErrMsg
					If InStr(sMessage, prnErrMsg) > 0 OR InStr(sMessage, "Line Stop Mode") > 0 Then
						Window("IpsEngine").Dialog("IPS Engine Stop").WinButton("OK").Click
						MyReportEvent micPass, "Printer Disconnected Message", "Message: "& prnErrMsg & " found!"
					Else
						call MyMsgBox(sMessage, sStepName) 
					End If
				End If
			ElseIf InStr(sMessage, "DRO") > 0 Then
				If InStr(sMessage, "Recover Case Print") > 0 Then
					If Window("TIPS").WinButton("Recover Case Printer").Exist(5) Then
						Window("TIPS").WinButton("Recover Case Printer").Click
						Wait 1
						If Dialog("InvalidNumOfArgs").Exist(2) Then
							Dialog("InvalidNumOfArgs").WinButton("OK").Click
						End If
						Window("TIPS").WinButton("Start IPS").Click
						Wait 1
					Else
						call MyMsgBox(sMessage, sStepName)
					End If
				ElseIf InStr(sMessage, "Recover Bundle Print") > 0 Then	
					If Window("TIPS").WinButton("Recover Bundle Print").Exist(5) Then
						Window("TIPS").WinButton("Recover Bundle Print").Click
						Wait 1
						If Dialog("InvalidNumOfArgs").Exist(2) Then
							Dialog("InvalidNumOfArgs").WinButton("OK").Click
						End If
						Window("TIPS").WinButton("Start IPS").Click
						Wait 1
					Else
						call MyMsgBox(sMessage, sStepName)
					End If
				ElseIf InStr(sMessage, "Recover Pallet Print") > 0 Then	
					If Window("TIPS").WinButton("Recover Pallet Print").Exist(5) Then
						Window("TIPS").WinButton("Recover Pallet Print").Click
						Wait 1
						If Dialog("InvalidNumOfArgs").Exist(2) Then
							Dialog("InvalidNumOfArgs").WinButton("OK").Click
						End If
						Window("TIPS").WinButton("Start IPS").Click
						Wait 1
					Else
						call MyMsgBox(sMessage, sStepName)
					End If
				ElseIf InStr(sMessage, "Recover Item Print") > 0 Then
					If Window("TIPS").WinButton("Recover Item Printer").Exist(5) Then
						Window("TIPS").WinButton("Recover Item Printer").Click
						Wait 1
						If Dialog("InvalidNumOfArgs").Exist(2) Then
							Dialog("InvalidNumOfArgs").WinButton("OK").Click
						End If
						Window("TIPS").WinButton("Start IPS").Click
						Wait 1
					Else
						call MyMsgBox(sMessage, sStepName)
					End If					
				ElseIf InStr(sMessage, "Recover Bundle Rework Print") > 0 Then	
					If Window("TIPS").WinButton("Rcvr Bundle RwPrint").Exist(5) Then
						Window("TIPS").WinButton("Rcvr Bundle RwPrint").Click
						Wait 1
						If Dialog("InvalidNumOfArgs").Exist(2) Then
							Dialog("InvalidNumOfArgs").WinButton("OK").Click
						End If
						Window("TIPS").WinButton("Start IPS").Click
						Wait 1
					Else
						call MyMsgBox(sMessage, sStepName)
					End If	
				ElseIf InStr(sMessage, "Recover Case Rework Print") > 0 Then
					If Window("TIPS").WinButton("Recvr Case RwPrinter").Exist(5) Then
						Window("TIPS").WinButton("Recvr Case RwPrinter").Click
						Wait 1
						If Dialog("InvalidNumOfArgs").Exist(2) Then
							Dialog("InvalidNumOfArgs").WinButton("OK").Click
						End If
						Window("TIPS").WinButton("Start IPS").Click
						Wait 1
					Else
						call MyMsgBox(sMessage, sStepName)
					End If		
				ElseIf InStr(sMessage, "Recover Pallet Rework Print") > 0 Then
					If Window("TIPS").WinButton("Recvr PalletRw Print").Exist(5) Then
						Window("TIPS").WinButton("Recvr PalletRw Print").Click
						Wait 1
						If Dialog("InvalidNumOfArgs").Exist(2) Then
							Dialog("InvalidNumOfArgs").WinButton("OK").Click
						End If
						Window("TIPS").WinButton("Start IPS").Click
						Wait 1
					Else
						call MyMsgBox(sMessage, sStepName)
					End If
				ElseIf InStr(sMessage, "Recover HU Print") > 0 Then
					If Window("TIPS").WinButton("Recvr HU Print").Exist(5) Then
						Window("TIPS").WinButton("Recvr HU Print").Click
						Wait 1
						If Dialog("InvalidNumOfArgs").Exist(2) Then
							Dialog("InvalidNumOfArgs").WinButton("OK").Click
						End If
						Window("TIPS").WinButton("Start IPS").Click
						Wait 1
					Else
						call MyMsgBox(sMessage, sStepName)
					End If
				ElseIf InStr(sMessage, "Start IPS") > 0 Then
					If Window("TIPS").WinButton("Start IPS").Exist(5) Then
						Window("TIPS").WinButton("Start IPS").Click
						Wait 1
						If Dialog("InvalidNumOfArgs").Exist(2) Then
							Dialog("InvalidNumOfArgs").WinButton("OK").Click
							Wait 1
						End If
						
					Else
						call MyMsgBox(sMessage, sStepName)
					End If	
				End If

			else
			    ' display message box
			     call MyMsgBox(sMessage, sStepName) 
			End if ' end if options check
		End If ' end if len()
		
	End Sub
	
	Public Function MyMsgBox(sMessage, sTitle)
		Dim vMsg
		If (m_bNoMsgBox = false) or (InStr(sMessage, "!!!!") > 0)  Then
			sMessage = sMessage & vbCr & vbCr & "YES to continue." & vbCr & "NO to fail test step." & vbCr & "CANCEL to abort test."
			vMsg = msgBox (sMessage, vbYesNoCancel, sTitle)
			If vMsg = vbCancel Then
				MyReportEvent micFail, sStepName, "Test ABORTED by operator"
				ExitTest
			End If
			
			If vMsg = vbNo Then
				MyReportEvent micFail, sStepName, "Operator Marked Test Failed."
			End If
		else
			wait 2, 0 ' wait 2 seconds instead of displaying message box
			vMsg = vbOK
		End If
		
		MyMsgBox = vMsg 
		
	End function
			
	Public Sub IpsOpen(sDirectory, sStationFiles)
	
		Call IpsLotEnd(60000)
	
		sDirectory = trim(sDirectory)
		sStatonFiles = trim(sStationFiles)
		' sDirectory directory for stations files, leave blank for D:\tips\config
		' sStationFiles are ',' seperated ips files
        sMsg = VisionSim.IpsOpen(sDirectory, sStationFiles)
        If len(sMsg)>0 And Left(sMsg,1) <> "*" Then
			MyReportEvent micPass, "Ips Files Opened", sStationFiles
		Else
			MyReportEvent micFail, "File Open Failed", sMsg
			ExitTest
        End If
	End Sub
	
	Public Sub IpsLotStart(nTimeoutMs)
        sMsg = VisionSim.IpsLotStart(nTimeoutMs)
	    If len(sMsg)>0 And Left(sMsg,1) <> "*" Then
			Reporter.ReportEvent micPass, "Lot Start", ""
		Else
		    Wait 10
		    p = false
		    For nTry = 1 To 10 Step 1
		    	Wait 10
		    	sMsg = VisionSim.IpsLotStart(nTimeoutMs)
		    	If len(sMsg)>0 And Left(sMsg,1) <> "*" Then
		    		p = true
					Reporter.ReportEvent micPass, "Lot Start", ""
					Exit For
				End If
		    Next
		    If p=false Then
		    	Reporter.ReportEvent micFail, "Lot Start Failed", sMsg
				ExitTest
		    End If
			
	    End If
	End Sub
	
	Public Sub IpsDeviceRecovery(nTimeoutMs)
        sMsg = VisionSim.IpsDeviceRecovery(nTimeoutMs)
        If len(sMsg)>0 And Left(sMsg,1) <> "*" Then
			MyReportEvent micPass, "Device Recovery", ""
		Else
			MyReportEvent micFail, "Device Recovery Failed", sMsg
			ExitTest
        End If
	End Sub
	
	Public Sub IpsLotEnd(nTimeoutMs)
        sMsg = VisionSim.IpsLotEnd(nTimeoutMs)
        If len(sMsg)>0 And Left(sMsg,1) <> "*" Then
			MyReportEvent micPass, "Lot Ended", ""
		Else
			Dim passflag
			passflag = false
			For nTry = 1 To 30 Step 1
				Wait 10
				sMsg = VisionSim.IpsLotEnd(nTimeoutMs)
				If len(sMsg)>0 And Left(sMsg,1) <> "*" Then
					passflag = true
					MyReportEvent micPass, "Lot Ended", ""
					Exit For
				End If
			Next
			If passflag = false Then
				MyReportEvent micFail, "Lot End Failed", sMsg
				ExitTest
			End If
        End If
	End Sub
	
	Public Sub DevMgrStartLot(sProductName, sVariables)	
		sMsg = VisionSim.DevMgrStartLot(sProductName, sVariables)
        Print "Lot Start Reply=" & sMsg
        If (len(sMsg)>0 And Left(sMsg,1) <> "*") Or len(sMsg)=0 Then
			MyReportEvent micPass, "DevMgr Lot Start", sMsg
		Else
			MyReportEvent micFail, "DevMgr Lot Start Failed", sMsg
			ExitTest
        End If
        wait(5)
        
    	' poll status for end of process
		' OpStatusIdle = 20,
		' OpStatusProcessing = 21,
		' OpStatusCanRetry = 22,
		' OpStatusFailed = 23,
		' OpStatusComplete = 24
        Dim sStatus 
        For nTry = 1 To 100
	        sStatus = VisionSim.DevMgrStatusCode()
			If sStatus <> 21 Then	' !processing
				Exit For
			End If
            Wait 10, 0
		Next

		If sStatus = 24 Then	' completed
			MyReportEvent micPass, "DevMgr Lot Start Completed", VisionSim.DevMgrStatusMessage()
		Else					' failed
			MyReportEvent micFail, "DevMgr Lot Start Incomplete", VisionSim.DevMgrStatusMessage()
			ExitTest
		End If
	End Sub

	Public Sub DevMgrStopLot()
		sMsg = VisionSim.DevMgrStopLot()
        Print "Lot End Reply=" & sMsg
        If len(sMsg)>0 And Left(sMsg,1) <> "*" Then
			MyReportEvent micPass, "DevMgr Lot End", ""
		Else
			MyReportEvent micFail, "DevMgr Lot End Failed", sMsg
			ExitTest
        End If
        wait(5)

    	' poll status for end of process
        Dim sStatus 
        For nTry = 1 To 100
	        sStatus = VisionSim.DevMgrStatusCode()
			If sStatus <> 21 Then	' !processing
				Exit For
			End If
            Wait 10, 0
		Next

		If sStatus = 24 Then	' completed
			MyReportEvent micPass, "DevMgr Lot End Completed", VisionSim.DevMgrStatusMessage()
		Else					' failed
			MyReportEvent micFail, "DevMgr Lot End Incomplete", VisionSim.DevMgrStatusMessage()
			ExitTest
		End If
	End Sub

	Public Sub DevMgrSuspendLot()
        sMsg = VisionSim.DevMgrSuspendLot()
        Print sMsg
        If len(sMsg)>0 And Left(sMsg,1) <> "*" Then
			MyReportEvent micPass, "DevMgr Lot Suspend", sMsg
		Else
			MyReportEvent micFail, "DevMgr Lot Suspend Failed", sMsg
			ExitTest
        End If

    	' poll status for end of process
        Dim sStatus 
        For nTry = 1 To 100
	        sStatus = VisionSim.DevMgrStatusCode()
			If sStatus <> 21 Then	' !processing
				Exit For
			End If
            Wait 10, 0
		Next

		If sStatus = 24 Then	' completed
			MyReportEvent micPass, "DevMgr Lot Suspend Completed", VisionSim.DevMgrStatusMessage()
		Else					' failed
			MyReportEvent micFail, "DevMgr Lot Suspend Incomplete", VisionSim.DevMgrStatusMessage()
			ExitTest
		End If
	End Sub

	Public Sub DevMgrResumeLot(sLotInfo)
        sMsg = VisionSim.DevMgrResumeLot(sLotInfo)
        Print sMsg
        If len(sMsg)>0 And Left(sMsg,1) <> "*" Then
			MyReportEvent micPass, "DevMgr Lot Resume", sMsg
		Else
			MyReportEvent micFail, "DevMgr Lot Resume Failed", sMsg
			ExitTest
        End If
        
    	' poll status for end of process
        Dim sStatus 
        For nTry = 1 To 100
	        sStatus = VisionSim.DevMgrStatusCode()
			If sStatus <> 21 Then	' !processing
				Exit For
			End If
            Wait 10, 0
		Next

		If sStatus = 24 Then	' completed
			MyReportEvent micPass, "DevMgr Lot Resume Completed", VisionSim.DevMgrStatusMessage()
		Else					' failed
			MyReportEvent micFail, "DevMgr Lot Resume Incomplete", VisionSim.DevMgrStatusMessage()
			ExitTest
		End If
	End Sub
	
End Class
 
Function Minapp
Set     qtApp = getObject("","QuickTest.Application")
    qtApp.WindowState = "Minimized"
    Set qtApp = Nothing

	
End Function

 'Printer Simulator Functions
 
