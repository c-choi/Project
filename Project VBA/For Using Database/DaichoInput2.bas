Attribute VB_Name = "DaichoInput"
Sub DaichoInput()

Dim Daicho As Workbook
Dim DbodyC As Integer
Dim DRowEnd As Range, DRowEndD As Range, DrowEndS As Range
Dim MbodyNum As Range
Dim SheetC As Integer
Dim sheetN As String
Dim Master As Workbook
Dim dateT As Date
Dim SheetD As Worksheet
Dim CountDumpW As Range, CountDumpS As Range, countdumpF As Range
Dim strAddW As String, strAddS As String, strAddF As String
Dim DumpRng As Range


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
Range(DRowEndD, DRowEndD.End(xlDown).End(xlDown).End(xlDown).Offset(0, 10)).Select
Selection.Clear

Master.Sheets(1).Activate
Z = 0
Set CountDumpW = MbodyNum.Offset(0, 11).Find(what:="ワイズダンプ", lookat:=xlPart)

If Not CountDumpW Is Nothing Then
strAddW = CountDumpW.Address

Do
w = w + 1

Set CountDumpW = MbodyNum.Offset(0, 11).FindNext(CountDumpW)
DRowEndD.Offset(Z, 5).Value = CountDumpW.Offset(0, -11).Value
Z = Z + 1
Loop While Not CountDumpW Is Nothing And strAddW <> CountDumpW.Address
End If
Set CountDumpW = Nothing

SheetD.Activate
Range(Range("a6"), Range("a6").End(xlToRight)).Copy

Range("a7").Offset(w + 2, 1).Select
Selection.Value = "セブン　保有車両"
Selection.Offset(1, -1).Select
Set DrowEndS = Selection
DrowEndS.PasteSpecial xlPasteAll

y = 0
Set CountDumpS = MbodyNum.Offset(0, 11).Find(what:="セブンダンプ", lookat:=xlPart)

If Not CountDumpS Is Nothing Then
strAddS = CountDumpS.Address

Do
s = s + 1

Set CountDumpS = MbodyNum.Offset(0, 11).FindNext(CountDumpS)
DrowEndS.Offset(y + 1, 5).Value = CountDumpS.Offset(0, -11).Value
y = y + 1
Loop While Not CountDumpS Is Nothing And strAddS <> CountDumpS.Address
End If
Set CountDumpS = Nothing

SheetD.Activate
Range(Range("a6"), Range("a6").End(xlToRight)).Copy

DrowEndS.Offset(s + 2, 1).Select
Selection.Value = "ホイ-ルクレ-ン"
Selection.Offset(1, -1).Select
Selection.PasteSpecial xlPasteAll

f = 0
Set countdumpF = MbodyNum.Offset(0, 11).Find(what:="ホイ-ルクレ-ン", lookat:=xlPart)

If Not countdumpF Is Nothing Then
strAddF = countdumpF.Address

Do
t = t + 1

Set countdumpF = MbodyNum.Offset(0, 11).FindNext(countdumpF)
DrowEndS.Offset(s + 2 + t + 1, 5).Value = countdumpF.Offset(0, -11).Value
y = y + 1
Loop While Not countdumpF Is Nothing And strAddF <> countdumpF.Address
End If
Set countdumpF = Nothing

SheetD.Activate
Range(Range("a6"), Range("a6").End(xlToRight)).Copy




Do Until i = SheetC
Daicho.Activate
Sheets(i).Activate
sheetN = ActiveSheet.Name
Range("a7").Activate
Set DRowEnd = ActiveCell
Range(DRowEnd, DRowEnd.End(xlDown).Offset(0, 10)).Select
Selection.Clear

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

ElseIf InStr(MbodyNum.Cells(j).Offset(0, 11).Value, sheetN) Then
If InStr(MbodyNum.Cells(j).Offset(0, 11).Value, "ダンプ") Or _
MbodyNum.Cells(j).Offset(0, 11).Value = "ホイ-ルクレ-ン" Then
Set DumpRng = SheetD.UsedRange.Find(what:=MbodyNum.Cells(j), lookat:=xlWhole)
MbodyNum.Cells(j).Offset(0, -4).Copy
DRowEnd.Offset(0, 1).PasteSpecial xlPasteValues
SheetD.Activate
DumpRng.Offset(0, -4).PasteSpecial xlPasteValues
Range(MbodyNum.Cells(j).Offset(0, -3), MbodyNum.Cells(j)).Copy
DRowEnd.Offset(0, 2).PasteSpecial xlPasteAll
DumpRng.Offset(0, -3).PasteSpecial xlPasteAll
Range(MbodyNum.Cells(j).Offset(0, 1), MbodyNum.Cells(j).Offset(0, 2)).Copy
DRowEnd.Offset(0, 7).PasteSpecial xlPasteAll
DumpRng.Offset(0, 2).PasteSpecial xlPasteAll
MbodyNum.Cells(j).Offset(0, 8).Copy
DRowEnd.Offset(0, 6).PasteSpecial xlPasteAll
DumpRng.Offset(0, 1).PasteSpecial xlPasteAll
Range(MbodyNum.Cells(j).Offset(0, 9), MbodyNum.Cells(j).Offset(0, 10)).Copy
DRowEnd.Offset(0, 9).PasteSpecial xlPasteAll
DumpRng.Offset(0, 4).PasteSpecial xlPasteAll
Daicho.Sheets(i).Activate
DRowEnd.Value = Application.WorksheetFunction.CountA(DRowEnd, Range(Range("a7"), Range("a7").End(xlDown))) + 1
Set DRowEnd = DRowEnd.Offset(1, 0)
SheetD.Activate
DumpRng.Offset(0, -5).Value = Application.WorksheetFunction.CountA(Range(DumpRng.End(xlUp).Offset(1, 0), DumpRng))
Set DRowEndD = DRowEndD.Offset(1, 0)

Else
MbodyNum.Cells(j).Offset(0, -4).Copy
DRowEnd.Offset(0, 1).PasteSpecial xlPasteValues
Range(MbodyNum.Cells(j).Offset(0, -3), MbodyNum.Cells(j)).Copy
DRowEnd.Offset(0, 2).PasteSpecial xlPasteAll
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

ElseIf InStr(MbodyNum.Cells(j).Offset(0, 11).Value, "ダンプ") Or _
MbodyNum.Cells(j).Offset(0, 11).Value = "ホイ-ルクレ-ン" Then
Set DumpRng = SheetD.UsedRange.Find(what:=MbodyNum.Cells(j), lookat:=xlWhole)
MbodyNum.Cells(j).Offset(0, -4).Copy
SheetD.Activate
DumpRng.Offset(0, -4).PasteSpecial xlPasteValues
Range(MbodyNum.Cells(j).Offset(0, -3), MbodyNum.Cells(j)).Copy
DumpRng.Offset(0, -3).PasteSpecial xlPasteAll
Range(MbodyNum.Cells(j).Offset(0, 1), MbodyNum.Cells(j).Offset(0, 2)).Copy
DumpRng.Offset(0, 2).PasteSpecial xlPasteAll
MbodyNum.Cells(j).Offset(0, 8).Copy
DumpRng.Offset(0, 1).PasteSpecial xlPasteAll
Range(MbodyNum.Cells(j).Offset(0, 9), MbodyNum.Cells(j).Offset(0, 10)).Copy
DumpRng.Offset(0, 4).PasteSpecial xlPasteAll
SheetD.Activate
DumpRng.Offset(0, -5).Value = Application.WorksheetFunction.CountA(Range(DumpRng.End(xlUp).Offset(1, 0), DumpRng))
Set DRowEndD = DRowEndD.Offset(1, 0)

ElseIf InStr(MbodyNum.Cells(j).Offset(0, 11).Value, "ダンプ") Or _
MbodyNum.Cells(j).Offset(0, 11).Value = "ホイ-ルクレ-ン" Then
Set DumpRng = SheetD.UsedRange.Find(what:=MbodyNum.Cells(j), lookat:=xlWhole)
MbodyNum.Cells(j).Offset(0, -4).Copy
SheetD.Activate
DumpRng.Offset(0, -4).PasteSpecial xlPasteValues
Range(MbodyNum.Cells(j).Offset(0, -3), MbodyNum.Cells(j)).Copy
DumpRng.Offset(0, -3).PasteSpecial xlPasteAll
Range(MbodyNum.Cells(j).Offset(0, 1), MbodyNum.Cells(j).Offset(0, 2)).Copy
DumpRng.Offset(0, 2).PasteSpecial xlPasteAll
MbodyNum.Cells(j).Offset(0, 8).Copy
DumpRng.Offset(0, 1).PasteSpecial xlPasteAll
Range(MbodyNum.Cells(j).Offset(0, 9), MbodyNum.Cells(j).Offset(0, 10)).Copy
DumpRng.Offset(0, 4).PasteSpecial xlPasteAll
SheetD.Activate
DumpRng.Offset(0, -5).Value = Application.WorksheetFunction.CountA(Range(DumpRng.End(xlUp).Offset(1, 0), DumpRng))
Set DRowEndD = DRowEndD.Offset(1, 0)
End If


j = j + 1

Loop

i = i + 1

Loop


Daicho.Activate

k = 1
For k = 1 To SheetC
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
If SheetD Then
Range(Range("a7").End(xlDown).End(xlDown).Offset(1, 0), Range("a7").End(xlDown).End(xlDown).Offset(1, 0).End(xlDown)).Select

Else
Range("d3").Value = Range("a7").End(xlDown).Value & "台"
End If
Range("a7").Activate
Next k

End If
Application.ScreenUpdating = True
Daicho.Sheets(1).Activate
'''ThisWorkbook.SaveCopyAs (Replace(dateT, "/", "") & "車両台帳 全体.xlsm")


End Sub
