Option Explicit

' PRE-REQUISITE: ERP Simulator must be running on Guardian server with custom WSEPCConfig.xml and option 'Do not change start number' enabled
' PRE-REQUISITE: UFT must be running (otherwise uncomment Launch command below)

Dim strStatus : strStatus = "Passed"
' Dim bRunning : bRunning = False
Dim strRoot : strRoot = "C:\Automation\Duplicate Check"
Dim qtApp 'As QuickTest.Application ' Declare the Application object variable
Set qtApp = CreateObject("QuickTest.Application") ' Create the Application object
qtApp.Launch ' Start UFT
qtApp.Visible = False ' Make the UFT application visible



''' Setup Environment '''''''''''
strStatus = RunAdvisor("SetupEnv")



WScript.echo "DONE @ " & Now() ' FOR TESTING PURPOSES ONLY

Set qtApp = Nothing ' Release the Application object


' DESC: Run Test on Advisor
'  xlsFile = Path of the file containing Action parameters for the Test
' RETURN: Status of the Test result (Passed or Failed or ?)
' NOTE: Does not log any info to result report
Function RunAdvisor(ByVal strTest)
	Dim strAdvisorTest : strAdvisorTest = strRoot & "\" &  strTest
	
	RunAdvisor = "NA"
	qtApp.Open strAdvisorTest, True ' Open in read-only mode

	qtApp.Test.Run Null, True ' Run the test and wait for completion
	RunAdvisor = qtApp.Test.LastRunResults.Status 
End Function