WorkingDirectory = "C:\Users\paul.flanagan\MainDirectory\CodeScraps\VBS\ExcelParsing\"
Set XLObj = CreateObject("Excel.Application")
Set WBObj = XLObj.Workbooks.Open(WorkingDirectory & "ParsingTargetFile.xlsx")
Set WSObj = WBObj.Sheets("Main")
Result1 = WSObj.Cells(1,1).value
Result2 = WSObj.Cells(2,1).value
Result3 = WSObj.Cells(3,1).value
Set XLObj = Nothing
Set XBObj = Nothing
Set XSObj = Nothing

msgbox(Result1)
msgbox(Result2)
msgbox(Result3)

Result2 = Split(Result2, ":")(1)
Result3 = Split(Result3, "(separator)")(1)

msgbox(Result2)
msgbox(Result3)

Set XLObj = CreateObject("Excel.Application")
Set WBObj = XLObj.Workbooks.Open(WorkingDirectory & "ParsingTargetFile.xlsx")
Set WSObj = WBObj.Sheets("Main")


''''''''''''''''
'    If WSObj.Cells(1,8).value <= 0 Then 
'		FileOpenObj.Writeline "successpercentage=" &FormatPercent(WSObj.Cells(1,8).value / WSObj.Cells(4,8).value, 0)
'		FileOpenObj.Write "genericsubjecttext=(" &Split(WSObj.Cells(2,3).value, "Advisor ")(1) &" Automation failed, " &FormatPercent(WSObj.Cells(1,8).value / WSObj.Cells(4,8).value, 0) &" successful)"
'	Else
'		FileOpenObj.Writeline "successpercentage=" &FormatPercent(WSObj.Cells(1,8).value / WSObj.Cells(4,8).value, 0)
'		FileOpenObj.Write "genericsubjecttext=(" &Split(WSObj.Cells(2,3).value, "Advisor ")(1) &" Automation " &FormatPercent(WSObj.Cells(1,8).value / WSObj.Cells(4,8).value, 0) &" successful)"
'	End If