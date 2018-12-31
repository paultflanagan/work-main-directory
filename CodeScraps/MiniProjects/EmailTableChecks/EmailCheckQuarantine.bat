@echo off
cls

>>C:\EmailTableChecks\EmailCheckCumulative.txt (
	type MostRecentLot.txt
	echo ****EmailCheckQuarantine.bat:****
	sqlcmd -S DupeServer -U sa -P cactus#1 -i C:\EmailTableChecks\EmailCheckAll.sql
	echo .
	echo .
	echo .
)

>C:\EmailTableChecks\EmailCheckQuarantineResults.txt (
	sqlcmd -S DupeServer -U sa -P cactus#1 -i C:\EmailTableChecks\EmailCheckQuarantine.sql
)