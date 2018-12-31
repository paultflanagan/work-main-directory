@echo off
cls

>>C:\EmailTableChecks\EmailCheckCumulative.txt (
	type MostRecentLot.txt
	echo ****EmailCheckAll.bat:****
	sqlcmd -S DupeServer -U sa -P cactus#1 -i C:\EmailTableChecks\EmailCheckAll.sql
	echo .
	echo .
	echo .
)

>C:\EmailTableChecks\EmailCheckAllResults.txt (
	sqlcmd -S DupeServer -U sa -P cactus#1 -i C:\EmailTableChecks\EmailCheckAll.sql
)