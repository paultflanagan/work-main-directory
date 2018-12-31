''Removing any lingering output files from previous runs
'msgbox("Attempting to remove any lingering output files from previous runs...")
'DeleteFile("C:\Automation\Duplicate Check\Results.txt")
'DeleteFile("C:\Automation\ReportFramework\Test_Results\Results.txt")
'DeleteFile("C:\Automation\ReportFramework\Test_Results\Project_Report.txt")
'DeleteFile("C:\Automation\ReportFramework\Test_Results\Project_Report_duplicatecheck.txt")
'DeleteFile("C:\3rdSafeguard.txt")
'msgbox("File removal attempt completed.")
'
''Creating new blank Results.txt file for Log output, since apparently the scripts do not do that automatically
'
'msgbox("Attempting to create blank Results.txt file...")
'Dim fso
'Set fso = CreateObject("Scripting.FileSystemObject") 
'
'fso.CreateTextFile("C:\Automation\Duplicate Check\Results.txt")
'
'Set fso = Nothing
'msgbox("File creation attempt completed.")
