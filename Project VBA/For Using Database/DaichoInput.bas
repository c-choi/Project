Attribute VB_Name = "DaichoInput"
Sub DaichoInput()

Dim Daicho As Workbook
Dim DbodyC As Integer
Dim DRowEnd As Range, DRowEndD As Range
Dim MbodyNum As Range
Dim SheetC As Integer
Dim sheetN As String
Dim Master As Workbook
Dim dateT As Date
Dim SheetD As Worksheet

dateT = Format(Now, YYYY - MM - DD)
Set Daicho = ThisWorkbook
Application.DisplayAlerts = False
On Error Resume Next
Set Master = Workbooks("ワイズ・セブンマスタファイル.xlsm")
 If Err.Number <> 0 Then
Filename = Application.GetOpenFilename
    If Filename = False Then
    Exit Sub
Else
Workbooks.Open Filename:=Filename, ReadOnly:=True
End If
Else
Master.Sheets(1).Activate
Set MbodyNum = Range(Range("h2"), Range("h2").End(xlDown))
Application.ScreenUpdating = False
Daicho.Activate
i = 1
j = 1
SheetC = Sheets.Count
Set SheetD = Daicho.Sheets("ダンプ保有一覧")
SheetD.Activate
SheetD.Range("a7").Activate
Set DRowEndD = ActiveCell
Range(DRowEndD, DRowEndD.End(xlDown).Offset(0, 10)).Select
Selection.Clear
Do Until i = SheetC
Daicho.Activate
Sheets(i).Activate
sheetN = ActiveSheet.Name
Range("a7").Activate
Set DRowEnd = ActiveCell
Range(DRowEnd, DRowEnd.End(xlDown).Offset(0, 10)).Select
Selection.Clear


Master.Sheets(1).Activate
j = 1

Do While j < MbodyNum.Count + 1

If MbodyNum.Cells(j).Offset(0, 11) = sheetN Then
MbodyNum.Cells(j).Offset(0, -4).Copy
DRowEnd.Offset(0, 1).PasteSpecial xlPasteValues
Range(MbodyNum.Cells(j).Offset(0, -3), MbodyNum.Cells(j)).Copy
DRowEnd.Offset(0, 2).PasteSpecial xlPasteAll
Master.Sheets(1).Activate
Range(MbodyNum.Cells(j).Offset(0, 1), MbodyNum.Cells(j).Offset(0, 2)).Copy
DRowEnd.Offset(0, 7).PasteSpecial xlPasteAll
MbodyNum.Cells(j).Offset(0, 8).Copy
DRowEnd.Offset(0, 6).PasteSpecial xlPasteAll
Range(MbodyNum.Cells(j).Offset(0, 9), MbodyNum.Cells(j).Offset(0, 10)).Copy
DRowEnd.Offset(0, 9).PasteSpecial xlPasteAll
Daicho.Sheets(i).Activate
DRowEnd.Value = Application.WorksheetFunction.CountA(DRowEnd, Range(Range("a7"), Range("a7").End(xlDown))) + 1
Set DRowEnd = DRowEnd.Offset(1, 0)

Else
If InStr(MbodyNum.Cells(j).Offset(0, 11).Value, sheetN) Then
If InStr(MbodyNum.Cells(j).Offset(0, 11).Value, "ダンプ") Then
MbodyNum.Cells(j).Offset(0, -4).Copy
DRowEnd.Offset(0, 1).PasteSpecial xlPasteValues
SheetD.Activate
DRowEndD.Offset(0, 1).PasteSpecial xlPasteValues
Range(MbodyNum.Cells(j).Offset(0, -3), MbodyNum.Cells(j)).Copy
DRowEnd.Offset(0, 2).PasteSpecial xlPasteAll
DRowEndD.Offset(0, 2).PasteSpecial xlPasteAll
Range(MbodyNum.Cells(j).Offset(0, 1), MbodyNum.Cells(j).Offset(0, 2)).Copy
DRowEnd.Offset(0, 7).PasteSpecial xlPasteAll
DRowEndD.Offset(0, 7).PasteSpecial xlPasteAll
MbodyNum.Cells(j).Offset(0, 8).Copy
DRowEnd.Offset(0, 6).PasteSpecial xlPasteAll
DRowEndD.Offset(0, 6).PasteSpecial xlPasteAll
Range(MbodyNum.Cells(j).Offset(0, 9), MbodyNum.Cells(j).Offset(0, 10)).Copy
DRowEnd.Offset(0, 9).PasteSpecial xlPasteAll
DRowEndD.Offset(0, 9).PasteSpecial xlPasteAll
Daicho.Sheets(i).Activate
DRowEnd.Value = Application.WorksheetFunction.CountA(DRowEnd, Range(Range("a7"), Range("a7").End(xlDown))) + 1
Set DRowEnd = DRowEnd.Offset(1, 0)
SheetD.Activate
DRowEndD.Value = Application.WorksheetFunction.CountA(DRowEndD, Range(Range("a7"), Range("a7").End(xlDown))) + 1
Set DRowEndD = DRowEndD.Offset(1, 0)

Else

MbodyNum.Cells(j).Offset(0, -4).Copy
DRowEnd.Offset(0, 1).PasteSpecial xlPasteValues
Range(MbodyNum.Cells(j).Offset(0, -3), MbodyNum.Cells(j)).Copy
DRowEnd.Offset(0, 2).PasteSpecial xlPasteAll
Master.Sheets(1).Activate
Range(MbodyNum.Cells(j).Offset(0, 1), MbodyNum.Cells(j).Offset(0, 2)).Copy
DRowEnd.Offset(0, 7).PasteSpecial xlPasteAll
MbodyNum.Cells(j).Offset(0, 8).Copy
DRowEnd.Offset(0, 6).PasteSpecial xlPasteAll
Range(MbodyNum.Cells(j).Offset(0, 9), MbodyNum.Cells(j).Offset(0, 10)).Copy
DRowEnd.Offset(0, 9).PasteSpecial xlPasteAll
Daicho.Sheets(i).Activate
DRowEnd.Value = Application.WorksheetFunction.CountA(DRowEnd, Range(Range("a7"), Range("a7").End(xlDown))) + 1
Set DRowEnd = DRowEnd.Offset(1, 0)
End If
ElseIf InStr(MbodyNum.Cells(j).Offset(0, 11).Value, "ダンプ") Then
MbodyNum.Cells(j).Offset(0, -4).Copy
SheetD.Activate
DRowEndD.Offset(0, 1).PasteSpecial xlPasteValues
Range(MbodyNum.Cells(j).Offset(0, -3), MbodyNum.Cells(j)).Copy
DRowEndD.Offset(0, 2).PasteSpecial xlPasteAll
Range(MbodyNum.Cells(j).Offset(0, 1), MbodyNum.Cells(j).Offset(0, 2)).Copy
DRowEndD.Offset(0, 7).PasteSpecial xlPasteAll
MbodyNum.Cells(j).Offset(0, 8).Copy
DRowEndD.Offset(0, 6).PasteSpecial xlPasteAll
Range(MbodyNum.Cells(j).Offset(0, 9), MbodyNum.Cells(j).Offset(0, 10)).Copy
DRowEndD.Offset(0, 9).PasteSpecial xlPasteAll
SheetD.Activate
DRowEndD.Value = Application.WorksheetFunction.CountA(DRowEndD, Range(Range("a7"), Range("a7").End(xlDown))) + 1
Set DRowEndD = DRowEndD.Offset(1, 0)

End If

End If
j = j + 1

Loop
i = i + 1

Loop

End If

Daicho.Activate
SheetD.Activate
DRowEndD.CurrentRegion.RemoveDuplicates Columns:=Array(2, 3, 4, 5, 6, 7, _
        8, 9, 10, 11), Header:=xlYes
k = 1
For k = 1 To SheetC + 1
ThisWorkbook.Sheets(k).Activate
Range(Range("a7"), Range("a7").End(xlDown).Offset(0, 10)).Select
With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
Range(Selection.End(xlDown).Offset(1, 0), Selection.End(xlDown).Offset(1, 10)).Select
Selection.ClearFormats
Range("a7").Activate
Next k
Application.ScreenUpdating = True
Daicho.Sheets(1).Activate
'''ThisWorkbook.SaveCopyAs (Replace(dateT, "/", "") & "車両台帳 全体.xlsm")


End Sub
