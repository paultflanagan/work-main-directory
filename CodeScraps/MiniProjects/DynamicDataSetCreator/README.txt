Title:   LotDataHandler
Version: 1.0
Author:  Paul F.
Date:    06-Nov-2018

Changelog:
    Date            Author          Description
    06-Nov-2018     Paul F.         Initial Version, limited to only accessing my server, DupeServer
    07-Nov-2018     Paul F.         Added support for connection to other servers, Settings.config, and README.txt


Summary:
    Connects to a SQL server, lists lots run since last data purge, and pulls complete data sets from desired lots.
    By modifying Settings.config, different Servers can be selected and credentials can be set.

Requirements:
    Access to the SQLCMD command, obtainable through the Microsoft SQL server Developer Edition
        Download-able at https://www.microsoft.com/en-us/sql-server/sql-server-downloads
    Storage of all required files in the same directory
    
Output:
    As many "Lot[#]Items.txt" files as were requested


Procedure:
    Ensure that Settings.config contains your desired ServerName, UserName, and Password
        If none are set, script will prompt for one-time-use target and credentials
    Run LotDataHandler.py
    When prompted, enter LotID of desired Lot
        View .\RecentLots.txt for more information on each lot
    Contents of selected lot will be output into "Lot[#]Items.txt"
    Afterwards, when prompted with option to pull from more lots, enter 'y' to proceed or 'n' to exit
    Additionally, enter "exit" at the LotID prompt to exit the program.
    