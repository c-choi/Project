Attribute VB_Name = "PlateNumberDivide"
Sub PlateNumberDivide()
''' for dividing Plate number copied from Daicho into Words and Numbers

Dim PtNum As Range
Dim FrnNum As Range
Dim BckNum As Range
Dim ID As Range
Dim BodyNum As Range
Dim BodyCount As Integer

Set BodyNum = Range(Range("j2"), Range("j2").End(xlDown))
BodyCount = BodyNum.Count
Set PtNum = Range(Range("a2"), Range("a2").Offset(BodyCount, 0))
Set ID = Range(Range("b2"), Range("b2").Offset(BodyCount, 0))
Set FrnNum = Range(Range("c2"), Range("c2").Offset(BodyCount, 0))
Set backnum = Range(Range("d2"), Range("d2").Offset(BodyCount, 0))

i = 1

Do Until i = BodyCount + 1
If PtNum.Cells(i).Value <> "" Then

If Left(PtNum.Cells(i), 3) = "èKéuñÏ" Then
If IsNumeric(Mid(PtNum.Cells(i), 7, 1)) Then

FrnNum.Cells(i).Value = Left(PtNum.Cells(i), 6)
backnum.Cells(i).Value = Mid(PtNum.Cells(i), 7, Len(PtNum.Cells(i)) - Len(FrnNum.Cells(i)))
ID.Cells(i).Value = backnum.Cells(i).Value
Else
FrnNum.Cells(i).Value = Left(PtNum.Cells(i), 7)
backnum.Cells(i).Value = Mid(PtNum.Cells(i), 8, Len(PtNum.Cells(i)) - Len(FrnNum.Cells(i)))
ID.Cells(i).Value = backnum.Cells(i).Value
End If

i = i + 1
Else
If IsNumeric(Mid(PtNum.Cells(i), 6, 1)) Then
FrnNum.Cells(i).Value = Left(PtNum.Cells(i), 5)
backnum.Cells(i).Value = Mid(PtNum.Cells(i), 6, Len(PtNum.Cells(i)) - Len(FrnNum.Cells(i)))
ID.Cells(i).Value = backnum.Cells(i).Value
Else
FrnNum.Cells(i).Value = Left(PtNum.Cells(i), 6)
backnum.Cells(i).Value = Mid(PtNum.Cells(i), 7, Len(PtNum.Cells(i)) - Len(FrnNum.Cells(i)))
ID.Cells(i).Value = backnum.Cells(i).Value

End If
i = i + 1
End If


Else
If Left(PtNum.Cells(i).Offset(0, 4), 3) = "èKéuñÏ" Then
If IsNumeric(Mid(PtNum.Cells(i).Offset(0, 4), 7, 1)) Then

FrnNum.Cells(i).Value = Left(PtNum.Cells(i).Offset(0, 4), 6)
backnum.Cells(i).Value = Mid(PtNum.Cells(i).Offset(0, 4), 7, Len(PtNum.Cells(i).Offset(0, 4)) - Len(FrnNum.Cells(i)))
ID.Cells(i).Value = backnum.Cells(i).Value
Else
FrnNum.Cells(i).Value = Left(PtNum.Cells(i).Offset(0, 4), 7)
backnum.Cells(i).Value = Mid(PtNum.Cells(i).Offset(0, 4), 8, Len(PtNum.Cells(i).Offset(0, 4)) - Len(FrnNum.Cells(i)))
ID.Cells(i).Value = backnum.Cells(i).Value
End If

i = i + 1
Else
If IsNumeric(Mid(PtNum.Cells(i).Offset(0, 4), 6, 1)) Then
FrnNum.Cells(i).Value = Left(PtNum.Cells(i).Offset(0, 4), 5)
backnum.Cells(i).Value = Mid(PtNum.Cells(i).Offset(0, 4), 6, Len(PtNum.Cells(i).Offset(0, 4)) - Len(FrnNum.Cells(i)))
ID.Cells(i).Value = backnum.Cells(i).Value
Else
FrnNum.Cells(i).Value = Left(PtNum.Cells(i).Offset(0, 4), 6)
backnum.Cells(i).Value = Mid(PtNum.Cells(i).Offset(0, 4), 7, Len(PtNum.Cells(i).Offset(0, 4)) - Len(FrnNum.Cells(i)))
ID.Cells(i).Value = backnum.Cells(i).Value

End If
i = i + 1
End If
End If
Loop

End Sub
