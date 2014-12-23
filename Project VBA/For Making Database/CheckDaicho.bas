Attribute VB_Name = "CheckDaicho"
Sub checkdaicho()
''' after adding Daicho data to Master File, compared ptnum and bodynum between two sheets and colored data

Dim LPlateNum As Range, LBodyNum As Range
Dim DplateNum As Range, DBodynum As Range
Dim iNum As String, jnum As String

Sheets(1).Activate
Set LPlateNum = Sheets(1).Range(Range("d2"), Range("d2").End(xlDown))
Set LBodyNum = Sheets(1).Range(Range("i2"), Range("i2").End(xlDown))
Sheets(2).Activate
Set DplateNum = Sheets(2).Range(Range("b2"), Range("b2").End(xlDown))
Set DBodynum = Sheets(2).Range(Range("f2"), Range("f2").End(xlDown))
i = 1


Do Until i = LPlateNum.Count + 1
Sheets(1).Activate
iNum = LPlateNum.Cells(i).Value
On Error Resume Next
If DplateNum.Find(what:=iNum, lookat:=xlWhole) = False Then
LPlateNum.Cells(i).Interior.Color = RGB(0, 255, 0)
i = i + 1
Else
DplateNum.Find(what:=LPlateNum.Cells(i), lookat:=xlWhole).Interior.Color = RGB(255, 255, 0)
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

Sub checkdaicho2()
'''after adding Baikyaku data, repeated check to see all number changes
Dim Platenum As Range, DplateNum As Range
Dim OldNum As Range

Sheets(1).Activate


Set Platenum = Sheets(1).Range(Range("i2"), Range("i2").End(xlDown))
Set OldNum = Platenum.Offset(0, 10)
Sheets(2).Activate

Set DplateNum = Sheets(2).Range(Range("f2"), Range("f2").End(xlDown))

i = 1
For i = 1 To Platenum.Count + 1
Do Until j = DplateNum.Count + 1
If Platenum.Cells(i) = DplateNum.Cells(j) Then
DplateNum.Cells(j).Interior.Color = RGB(255, 255, 255)
If OldNum.Cells(i).Value = "" Then
OldNum.Cells(i).Value = "�ԔԕύXX"
j = 1
i = i + 1
OldNum.Cells(i).Interior.Color = RGB(255, 255, 255)
Else
j = 1
i = i + 1
End If
Else
OldNum.Cells(i).Interior.Color = RGB(255, 255, 255)
j = j + 1
End If
Loop
On Error Resume Next
j = 1
Next i

Sheets(1).Activate
End Sub
