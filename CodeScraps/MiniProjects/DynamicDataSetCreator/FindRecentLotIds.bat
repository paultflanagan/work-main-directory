@echo off
REM cls

>.\RecentLots.txt (
    sqlcmd -S %1 -U %2 -P %3 -i .\FindRecentLotIds.sql
)

>.\RecentLotIDs.txt (
    FOR /F "skip=3 tokens=1 eol=(" %%a IN (.\RecentLots.txt) DO (
        echo %%a
    )
)