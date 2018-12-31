'Set Working Directory
WD = "C:\Users\paul.flanagan\MainDirectory\DeleteFileWildTesting\Files"

CreateFile(WD & "\Alpha\a1.txt")
CreateFile(WD & "\Alpha\a2.txt")
CreateFile(WD & "\Alpha\a3.txt")
CreateFile(WD & "\Alpha\AAlpha\AAAlpha\aaa1.txt")
CreateFile(WD & "\Alpha\AAlpha\AAAlpha\aaa2.txt")
CreateFile(WD & "\Alpha\AAlpha\AABravo\aab1.txt")
CreateFile(WD & "\Alpha\ABravo\ab1.txt")
CreateFile(WD & "\Alpha\ABravo\ab2.txt")
CreateFile(WD & "\Bravo\b1.txt")
CreateFile(WD & "\Bravo\BAlpha\ba1.txt")
CreateFile(WD & "\Bravo\BBravo\bb1.txt")
CreateFile(WD & "\Bravo\BBravo\bb2.txt")
CreateFile(WD & "\Bravo\BBravo\bb3.txt")
CreateFile(WD & "\Bravo\BCharlie\bc1.txt")
CreateFile(WD & "\Charlie\ca.txt")
CreateFile(WD & "\Charlie\cb.txt")

Sub CreateFile(ByVal FileName)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	Call fso.CreateTextFile(FileName, True)
	'Print("Created File " & FileName)
	Set fso = Nothing
End Sub