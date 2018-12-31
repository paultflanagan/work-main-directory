@echo off
cls

>C:\2ndSafeguard.txt (
	sqlcmd -S DupeServer -U sa -P cactus#1 -i C:\Automation\sg2.sql
)
