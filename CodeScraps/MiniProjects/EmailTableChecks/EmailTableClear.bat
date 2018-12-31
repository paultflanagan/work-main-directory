@echo off

cls

sqlcmd -S DupeServer -U sa -P cactus#1 -i C:\EmailTableChecks\EmailTableClear.sql

>>C:\EmailTableChecks\EmailCheckCumulative.txt (
	echo ****Clearing Email Table****
	echo .
	echo .
	echo .
	echo .
	echo .
	echo .
)
