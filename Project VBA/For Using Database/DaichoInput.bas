Attribute VB_Name = "DaichoInput"
Sub DaichoInput()
Dim Daicho As ThisWorkbook
Dim RegNum As Range, FirstReg As Range, Model As Range, MakerName As Range
Dim BodyNum As Range, CarType As Range, MaxW As Range, TotalW As Range, NoxPM As Range, LevNum As Range
Dim SheetC As Integer

SheetC = Daicho.Sheets.Count
i = 1

Do While i = SheetC
Sheets(i).Activate

Set RegNum = Range("b6")
Set FirstReg = Range("c6")
Set Model = Range("d6")
Set MakerName = Range("e6")
Set BodyNum = Range("f6")
Set CarType = Range("g6")
Set MaxW = Range("h6")
Set TotalW = Range("i6")
Set NoxPM = Range("j6")
Set LevNum = Range("k6")




End Sub
