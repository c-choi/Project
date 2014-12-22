Attribute VB_Name = "CheckDaicho"
Sub checkdaicho()
''' after adding Daicho data to Master File, compared ptnum and bodynum between two sheets and colored data

Dim LPlateNum As Range, LBodyNum As Range
Dim DPlateNum As Range, DBodynum As Range
Dim iNum As String, jnum As String

Sheets(1).Activate
Set LPlateNum = Sheets(1).Range(Range("d2"), Range("d2").End(xlDown))
Set LBodyNum = Sheets(1).Range(Range("i2"), Range("i2").End(xlDown))
Sheets(2).Activate
Set DPlateNum = Sheets(2).Range(Range("b2"), Range("b2").End(xlDown))
Set DBodynum = Sheets(2).Range(Range("f2"), Range("f2").End(xlDown))
i = 1


Do Until i = LPlateNum.Count + 1
Sheets(1).Activate
iNum = LPlateNum.Cells(i).Value
On Error Resume Next
If DPlateNum.Find(what:=iNum, lookat:=xlWhole) = False Then
LPlateNum.Cells(i).Interior.Color = RGB(0, 255, 0)
i = i + 1
Else
DPlateNum.Find(what:=LPlateNum.Cells(i), lookat:=xlWhole).Interior.Color = RGB(255, 255, 0)
i = i + 1
End If
Loop

j = 1
Do Until j = LBodyNum.Count + 1
Sheets(1).Activate
jnum = LBodyNum.Cells(j).Value
On Error Resume Next
If DBodynum.Find(what:=jnum, lookat:=xlWhole) = False Then
LBodyNum.Cells(j).Interior.Color = RGB(0, 255, 0)
j = j + 1
Else
DBodynum.Find(what:=LBodyNum.Cells(j), lookat:=xlWhole).Interior.Color = RGB(255, 255, 0)
j = j + 1
End If
Loop

End Sub

