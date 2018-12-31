@echo off
cls

>>C:\EmailTableChecks\EmailCheckCumulative.txt (
	echo ****EmailCheckReuse.bat:****
	sqlcmd -S DupeServer -U sa -P cactus#1 -i C:\EmailTableChecks\EmailCheckAll.sql
	echo .
	echo .
	echo .
)

>C:\EmailTableChecks\EmailCheckReuseResults.txt (
	sqlcmd -S DupeServer -U sa -P cactus#1 -i C:\EmailTableChecks\EmailCheckReuse.sql
)