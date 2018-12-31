:: Environment Cleanup
:: Removing any lingering output files from previous runs
DEL "C:\Automation\Duplicate Check\Results.txt"
DEL "C:\Automation\ReportFramework\Test_Results\Results.txt"
DEL "C:\Automation\ReportFramework\Test_Results\Project_Report.txt"
DEL "C:\Automation\ReportFramework\Test_Results\Project_Report_duplicatecheck.txt"
DEL "C:\2ndSafeguard.txt"
DEL "C:\3rdSafeguard.txt"
echo Y | DEL "C:\Automation\Duplicate Check\Emails\*.*"