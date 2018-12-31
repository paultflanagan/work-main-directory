@echo off
cls

XCOPY "C:\PimLabTestData\DuplicateCheckPimLabTestData_3level_GoodLot_810_Standard.xlsx" "C:\PimLabTestData\DuplicateCheckPimLabTestData_3level_GoodLot_810.xlsx" /c /r /y
XCOPY "C:\PimLabTestData\DuplicateCheckPimLabTestData_3level_RemoveDupes_810_Standard.xlsx" "C:\PimLabTestData\DuplicateCheckPimLabTestData_3level_RemoveDupes_810.xlsx" /c /r /y
XCOPY "C:\PimLabTestData\DuplicateCheckPimLabTestData_3level_AcceptQuarantine_810_Standard.xlsx" "C:\PimLabTestData\DuplicateCheckPimLabTestData_3level_AcceptQuarantine_810.xlsx" /c /r /y
XCOPY "C:\PimLabTestData\DuplicateCheckPimLabTestData_3level_GoodLot_810_Standard.xlsx" "C:\PimLabTestData\DuplicateCheckPimLabTestData_3level_ExcludeDataname_810.xlsx" /c /r /y

XCOPY "C:\Automation\Duplicate Check\ImportFiles\Functional\Lot_Good\FT-B_List_Carton_10000_Standard.xml" "C:\Automation\Duplicate Check\ImportFiles\Functional\Lot_Good\FT-B_List_Carton_10000.xml" /c /r /y
XCOPY "C:\Automation\Duplicate Check\ImportFiles\Functional\Lot_RemoveDuplicates\FT-B_List_Carton_40000_Standard.xml" "C:\Automation\Duplicate Check\ImportFiles\Functional\Lot_RemoveDuplicates\FT-B_List_Carton_40000.xml" /c /r /y
XCOPY "C:\Automation\Duplicate Check\ImportFiles\Functional\Lot_QAccept\FT-B_List_Carton_70000_Standard.xml" "C:\Automation\Duplicate Check\ImportFiles\Functional\Lot_QAccept\FT-B_List_Carton_70000.xml" /c /r /y
XCOPY "C:\Automation\Duplicate Check\ImportFiles\Functional\Lot_ExcludeDataname\FT-B_List_Carton_160000_Standard.xml" "C:\Automation\Duplicate Check\ImportFiles\Functional\Lot_ExcludeDataname\FT-B_List_Carton_160000.xml" /c /r /y