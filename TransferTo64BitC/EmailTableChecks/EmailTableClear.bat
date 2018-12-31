@echo off

cls

sqlcmd -S DupeServer -U sa -P cactus#1 -i C:\EmailTableChecks\EmailTableClear.sql