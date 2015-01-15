Attribute VB_Name = "Insurance"
Sub Insurance()

Dim MCarNum As Range
Dim ICarNum As Range
Dim IcarPlate As String
Dim InsFile As Workbook
Dim Master As Workbook
Dim MasterList As Worksheet
Dim MInsNum As Range, IInsNum As Range
Dim RowCount As Integer
Dim FindNum As Range

Application.ScreenUpdating = False
Set InsFile = Workbooks("ïΩê¨26îN12åé20ì˙Å`Å@é©ìÆé‘ï€åØñæç◊èëM.xlsx")
Set Master = ThisWorkbook
Set MasterList = Master.Sheets("é‘óºàÍóó")

MasterList.Activate
Set MCarNum = MasterList.Range(Range("d2"), Range("d2").End(xlDown))
InsFile.Activate
Set ICarNum = InsFile.Sheets(1).Range(Range("c2"), Range("c2").End(xlDown))

RowCount = ICarNum.Cells.Count

For i = 1 To RowCount
IcarPlate = ICarNum.Cells(i).Value & ICarNum.Cells(i).Offset(0, 1).Value
Set FindNum = MCarNum.Find(what:=IcarPlate, LookIn:=xlValues, lookat:=xlWhole)
If FindNum Is Nothing = 0 Then
ICarNum.Cells(i).Offset(0, -1).Copy
FindNum.Offset(0, 33).PasteSpecial xlPasteValues
InsFile.Activate
ICarNum.Cells(i).Offset(0, 3).Select
Range(Selection, Selection.Offset(0, 1)).Select
Selection.Copy
FindNum.Offset(0, 34).PasteSpecial xlPasteValues
FindNum.Offset(0, 2).Copy
ICarNum.Cells(i).Offset(0, 2).PasteSpecial xlPasteValues
Else
End If
Next i

Application.ScreenUpdating = True
End Sub
