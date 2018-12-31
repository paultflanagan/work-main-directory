@echo off
cls

>C:\3rdSafeguard.txt (
	sqlcmd -S DupeServer -U sa -P cactus#1 -i C:\Automation\sg3.sql
)
