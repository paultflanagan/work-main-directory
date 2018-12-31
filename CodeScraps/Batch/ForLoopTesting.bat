@echo off
cls

FOR /F "skip=3 tokens=1 eol=(" %%a IN (.\ForLoopSource.txt) DO (
    echo %%a
)