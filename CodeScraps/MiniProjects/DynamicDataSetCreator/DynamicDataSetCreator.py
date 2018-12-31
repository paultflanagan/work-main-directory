# @Title:   LotDataHandler
# @Version: 1.0
# @Author:  Paul F.
# @Date:    06-Nov-2018
# Summary:
#   Connects to a SQL server, lists lots run since last data purge, and pulls complete data sets from desired lots.
#   By modifying Settings.config, different Servers can be selected and credentials can be set.
#
# Requirements:
#   Access to the SQLCMD command, obtainable through the Microsoft SQL server Developer Edition (see README.txt for details)
#   Storage of all required folders in the same directory
#
# Output:
#   As many "Lot[#]Items.txt" files as were requested
#
# Changelog:
#   Date            Author          Description
#   06-Nov-2018     Paul F.         Initial Version, limited to only accessing DupeServer
#   07-Nov-2018     Paul F.         Added support for connection to other servers, Settings.config, and README.txt
#

import subprocess
import sys
import io
import time

def ProgramExit():
    sys.exit()

print("Loading Configuration Variables from .\Settings.config...")

ServerName = ''
UserName = ''
Password = ''

Config_FileObject = open(".\Settings.config", "r")
SettingsList = Config_FileObject.read().split('\n')
Config_FileObject.close()
for setting in SettingsList:
    settingSplitPair = setting.split('=')
    if settingSplitPair[0] == "ServerName":
        ServerName = settingSplitPair[1]
    elif settingSplitPair[0] == "UserName":
        UserName = settingSplitPair[1]
    elif settingSplitPair[0] == "Password":
        Password = settingSplitPair[1]

if ServerName == '':
    print('No value found for "ServerName". You can edit Settings.config to set a default value for future use.')
    ServerName = input("What is the target ServerName? ")
if UserName == '':
    print('No value found for "UserName". You can edit Settings.config to set a default value for future use.')
    ServerName = input("What is the desired UserName? ")
if Password == '':
    print('No value found for "Password". You can edit Settings.config to set a default value for future use.')
    ServerName = input("What is the desired Password? ")


print("Finding recent Lots...")

subprocess.call("C:\Windows\System32\cmd.exe /C .\FindRecentLotIds.bat {0} {1} {2}".format(ServerName, UserName, Password))

#Read the file, take off any trailing whitespace (the final empty line), put into list delimited by new lines.
RecentLotIDs_FileObject = open(".\RecentLotIDs.txt", "r")
LotIDList = RecentLotIDs_FileObject.read().strip().split('\n')
RecentLotIDs_FileObject.close()

#if len(LotIDList) == 0:
if len(LotIDList) == 1 and LotIDList[0] == '':
    print("No recent Lots found.")
    ProgramExit()

moreLotsToProcess = True
while moreLotsToProcess:

    desiredLot = LotIDList[0]

    print("Exporting contents of desired lot to .\Lot{0}Items.txt...".format(desiredLot))
    subprocess.call("C:\Windows\System32\cmd.exe /C .\FindNumbersWithinLot {0} {1} {2} {3}".format(ServerName, UserName, Password, desiredLot))
    print("Contents exported.")
    validSelection = False
    while not validSelection:
        print("Would you like to process an additional lot?")
        response = input("(y/n): ")
        if response.lower() == "n":
            validSelection = True
            moreLotsToProcess = False
        elif response.lower() == "y":
            validSelection = True
        else:
            print("Invalid selection: please respond 'y' or 'n'.")

ProgramExit()
