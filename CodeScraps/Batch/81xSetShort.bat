@echo off
cls

XCOPY "C:\PimLabTestData\DuplicateCheckPimLabTestData_3level_GoodLot_810_Short.xlsx" "C:\PimLabTestData\DuplicateCheckPimLabTestData_3level_GoodLot_810.xlsx" /c /r /y
XCOPY "C:\PimLabTestData\DuplicateCheckPimLabTestData_3level_RemoveDupes_810_Short.xlsx" "C:\PimLabTestData\DuplicateCheckPimLabTestData_3level_RemoveDupes_810.xlsx" /c /r /y
XCOPY "C:\PimLabTestData\DuplicateCheckPimLabTestData_3level_AcceptQuarantine_810_Short.xlsx" "C:\PimLabTestData\DuplicateCheckPimLabTestData_3level_AcceptQuarantine_810.xlsx" /c /r /y
XCOPY "C:\PimLabTestData\DuplicateCheckPimLabTestData_3level_GoodLot_810_Short.xlsx" "C:\PimLabTestData\DuplicateCheckPimLabTestData_3level_ExcludeDataname_810.xlsx" /c /r /y

XCOPY "C:\Automation\Duplicate Check\ImportFiles\Functional\Lot_Good\FT-B_List_Carton_10000_Short.xml" "C:\Automation\Duplicate Check\ImportFiles\Functional\Lot_Good\FT-B_List_Carton_10000.xml" /c /r /y
XCOPY "C:\Automation\Duplicate Check\ImportFiles\Functional\Lot_RemoveDuplicates\FT-B_List_Carton_40000_Short.xml" "C:\Automation\Duplicate Check\ImportFiles\Functional\Lot_RemoveDuplicates\FT-B_List_Carton_40000.xml" /c /r /y
XCOPY "C:\Automation\Duplicate Check\ImportFiles\Functional\Lot_QAccept\FT-B_List_Carton_70000_Short.xml" "C:\Automation\Duplicate Check\ImportFiles\Functional\Lot_QAccept\FT-B_List_Carton_70000.xml" /c /r /y
XCOPY "C:\Automation\Duplicate Check\ImportFiles\Functional\Lot_ExcludeDataname\FT-B_List_Carton_160000_Short.xml" "C:\Automation\Duplicate Check\ImportFiles\Functional\Lot_ExcludeDataname\FT-B_List_Carton_160000.xml" /c /r /y