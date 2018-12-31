'------------------------------------------------------------------
'   Description   	  :      Loads the PIM Framework spreadsheet to be used for this test
'   Project           :      Uniseries Duplicate Check 7.40.011
'   Author            :      Paul F.
'   © 2018   Systech International.  All rights reserved
'------------------------------------------------------------------
'
'	Prologue:
'   - Product Data has been loaded through FT_DualFormat_Driver via Guardian
'   
'   Epilogue:
'   - Loads the PIM spreadsheet that will run the line.

call PIM.InitializeFromData()
