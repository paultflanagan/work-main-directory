@echo off
REM cls

>.\Lot%4Items.txt (
    sqlcmd -v DesiredID =%4 -S %1 -U %2 -P %3 -i .\FindNumbersWithinLot.sql
)