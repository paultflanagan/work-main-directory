'Open, start, and login to TIPS
Print("Attempting Diagnostic Screen Manager Login.")
totalTries = 0
While (NOT(Window("Menu").WinButton("Products").Exist(5)) AND totalTries < 2)
	boolFunc_StartTipsDiagnosticAndLogin "superuser", "master"
	counter = 0
	While (NOT(Window("Menu").WinButton("Products").Exist(0)) AND counter < 12)
		wait(10)
		counter = counter + 1
	Wend
	totalTries = totalTries + 1
Wend

If NOT(Window("Menu").WinButton("Products").Exist(0)) Then
	print("Screen Manager Login Failed")
End If

Print("Diagnostic Screen Manager Login Attempt Complete.")

'Check the State by TagValue, returns 0/1 based on the Tag value entered.
Function ReadTag(ByVal strTag)
    set rtdb = createobject("TIPS.RTDB.2")
    ReadTag = rtdb.ReadTagField(strTag, "A_CV")
    Set rtdb = nothing
 end function
 
'Login into tips 
Function boolFunc_StartTipsDiagnosticAndLogin(ByVal strUserId, ByVal strUserPwd)
boolFunc_StartTipsAndLogin = False
	
	'Open Advisor If Not Running
	If Not Window("Menu").Exist(0) Then
		Set oShell=CreateObject("WScript.Shell")
			oShell.run "cmd /C CD C:\ & menuman.exe /d",0,False
		Set oShell = nothing
	End If
	
	'Navigate the diagnostic menu window
	Window("Screen Manager Diagnostic").Activate
	Window("Screen Manager Diagnostic").Type(micAltDwn)
	Window("Screen Manager Diagnostic").Type(micAltUp)
	Window("Screen Manager Diagnostic").Type("d")
	Window("Screen Manager Diagnostic").Type("a")
	
	'Loop until advisor loads completely, if started.
	Do While Dialog("Screen Manager").Exist(2) = "True" 
		If Window("Menu").WinButton("Login").Exist(0)Then
			Exit do
		End If
	loop
	
	'Check if Login button is visible to click and login advisor
	If Window("Menu").WinButton("Login").WaitProperty("visible","true",900000) Then
		Window("Menu").WinButton("Login").Click

		With Dialog("User Login")
			'Set the value for username		
			'.WinEdit(editUserID).Set strUserId 	
			.WinEdit("User Id:").Set strUserId
			
			'Set the password value as secured and encrypted in datatable.
			'.WinEdit(editPwd).SetSecure strUserPwd 	
			.WinEdit("Password").SetSecure strUserPwd

			'Click ok to login		
			'.WinButton(btnOk).Click
			.WinButton("OK").Click

			'Check If Advisor application logged in, Function return true if Login successful.
			If Window("Menu").WinButton("Products").WaitProperty("visible","true",900000) AND ReadTag("LoggedIn") = 1 Then		
				boolFunc_StartTipsAndLogin = True
			End If		
		End With
	End If  
	Print("Completed Screen Manager Login function")
End Function
