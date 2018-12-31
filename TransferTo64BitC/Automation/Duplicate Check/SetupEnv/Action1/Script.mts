''' Environment Cleanup '''''''''
'Removing any lingering output files from previous runs
DeleteFile("C:\Automation\Duplicate Check\Results.txt")
DeleteFile("C:\Automation\ReportFramework\Test_Results\Results.txt")
DeleteFile("C:\Automation\ReportFramework\Test_Results\Project_Report.txt")
DeleteFile("C:\Automation\ReportFramework\Test_Results\Project_Report_duplicatecheck.txt")
DeleteFile("C:\3rdSafeguard.txt")
DeleteFile("C:\Automation\Duplicate Check\Emails\Email_1.txt")
DeleteFile("C:\Automation\Duplicate Check\Emails\Email_2.txt")
DeleteFile("C:\Automation\Duplicate Check\Emails\Email_3.txt")
DeleteFile("D:\temp\QTPrintLog.txt")
'DeleteFile("C:\Automation\ReportFramework\Test_Results\QTPrintLog.txt")
DeleteFile("C:\EmailCheckCumulative.txt")
DeleteFile("C:\3rdSafeguard_1.txt")
DeleteFile("C:\3rdSafeguard_2.txt")
DeleteFile("C:\3rdSafeguard_3.txt")
DeleteFile("C:\3rdSafeguard_4.txt")
DeleteFile("C:\3rdSafeguard_5.txt")
DeleteFile("C:\3rdSafeguard_6.txt")


Set Wshell = CreateObject("wscript.shell")

'Clear out SQL Email Table
Wshell.Run "C:\EmailTableClear.bat"

'Close Guardian process
Wshell.Run "taskkill /IM GuardianSPTConfig.EXE /F"

'Clear all TIPS processes
Wshell.Run "D:\Tips\Bin\TipsKillAllAuto.cmd", True
'While Window("C:\Windows\system32\cmd.exe").Exist(2)
'	wait(1)
'Wend
	
Set Wshell = nothing


