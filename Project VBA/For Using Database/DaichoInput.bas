Attribute VB_Name = "DaichoInput"
Sub DaichoInput()

Dim Daicho As Workbook
Dim DbodyC As Integer
Dim DRowEnd As Range
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
Workbooks.Open (Filename)
End If
Else
Master.Sheets(1).Activate
Set MbodyNum = Range(Range("h2"), Range("h2").End(xlDown))
Application.ScreenUpdating = False
Daicho.Activate
i = 1
j = 1
SheetC = Sheets.Count
SheetD = Daicho.Sheets("ダンプ保有一覧")

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

Do While j < MbodyNum.Count

If MbodyNum.Cells(j).Offset(0, 11) = sheetN Then
MbodyNum.Cells(j).Offset(0, -4).Copy
DRowEnd.Offset(0, 1).PasteSpecial xlPasteValues
Daicho.Sheets(i).Activate
Range(DRowEnd, DRowEnd.Offset(0, 1)).Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
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
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
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

ElseIf MbodyNum.Cells(j).Offset(0, 11) Like "[*ダンプ*]" Then
MbodyNum.Cells(j).Offset(0, -4).Copy
SheetD.Activate
Set DRowEnd = Range("a7")
DRowEnd.Offset(0, 1).PasteSpecial xlPasteValues
Range(DRowEnd, DRowEnd.Offset(0, 1)).Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
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
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
Range(MbodyNum.Cells(j).Offset(0, -3), MbodyNum.Cells(j)).Copy
DRowEnd.Offset(0, 2).PasteSpecial xlPasteAll
Master.Sheets(1).Activate
Range(MbodyNum.Cells(j).Offset(0, 1), MbodyNum.Cells(j).Offset(0, 2)).Copy
DRowEnd.Offset(0, 7).PasteSpecial xlPasteAll
MbodyNum.Cells(j).Offset(0, 8).Copy
DRowEnd.Offset(0, 6).PasteSpecial xlPasteAll
Range(MbodyNum.Cells(j).Offset(0, 9), MbodyNum.Cells(j).Offset(0, 10)).Copy
DRowEnd.Offset(0, 9).PasteSpecial xlPasteAll
SheetD.Activate
DRowEnd.Value = Application.WorksheetFunction.CountA(DRowEnd, Range(Range("a7"), Range("a7").End(xlDown))) + 1
Set DRowEnd = DRowEnd.Offset(1, 0)

End If
j = j + 1

Loop
i = i + 1
Loop
End If
Master.Activate
Application.ScreenUpdating = True
Daicho.Activate
'''ThisWorkbook.SaveCopyAs (Replace(dateT, "/", "") & "車両台帳 全体.xlsm")


End Sub
