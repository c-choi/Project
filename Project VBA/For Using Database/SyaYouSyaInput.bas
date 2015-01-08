Attribute VB_Name = "SyaYouSyaInput"
Sub SyaYoInput()

Dim SyaYo As Workbook
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
Dim strFP As String

strFP = ThisWorkbook.Path
dateT = Format(Now, YYYY - MM - DD)

Set SyaYo = ThisWorkbook
Application.DisplayAlerts = False
Application.ScreenUpdating = False
On Error Resume Next
Set Master = Workbooks("ワイズ・セブンマスタファイル.xlsm")
 If Err.Number <> 0 Then
Filename = Application.GetOpenFilename
    If Filename = False Then
    Exit Sub
Else
Workbooks.Open Filename:=Filename, ReadOnly:=True
SyaYoInput
End If
Else
Master.Sheets(1).Activate
Set MbodyNum = Range(Range("h2"), Range("h2").End(xlDown))
SyaYo.Activate
I = 1
j = 1

SheetC = Sheets.Count



Do Until I = SheetC
SyaYo.Activate
Sheets(I).Activate
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
MbodyNum.Cells(j).Offset(0, 3).Copy
DRowEnd.Offset(0, 7).PasteSpecial xlPasteAll
MbodyNum.Cells(j).Offset(0, 8).Copy
DRowEnd.Offset(0, 6).PasteSpecial xlPasteAll
MbodyNum.Cells(j).Offset(0, 17).Copy
DRowEnd.Offset(0, 8).PasteSpecial xlPasteAll
MbodyNum.Cells(j).Offset(0, 5).Copy
DRowEnd.Offset(0, 9).PasteSpecial xlPasteAll
MbodyNum.Cells(j).Offset(0, 4).Copy
DRowEnd.Offset(0, 10).PasteSpecial xlPasteAll
If MbodyNum.Cells(j).Offset(0, 20).Value = "X" Then
DRowEnd.Offset(0, 11).Value = "売却"
DRowEnd.Interior.Color = RGB(0, 0, 20)
SyaYo.Sheets(I).Activate
Set DRowEnd = DRowEnd.Offset(1, 0)
Else
SyaYo.Sheets(I).Activate
DRowEnd.Value = Application.WorksheetFunction.CountA(DRowEnd, Range(Range("a7"), Range("a7").End(xlDown))) + 1
Set DRowEnd = DRowEnd.Offset(1, 0)
End If
j = j + 1
Else
If InStr(MbodyNum.Cells(j).Offset(0, 11).Value, sheetN) Then
MbodyNum.Cells(j).Offset(0, -4).Copy
DRowEnd.Offset(0, 1).PasteSpecial xlPasteValues
Range(MbodyNum.Cells(j).Offset(0, -3), MbodyNum.Cells(j)).Copy
DRowEnd.Offset(0, 2).PasteSpecial xlPasteAll
Master.Sheets(1).Activate
MbodyNum.Cells(j).Offset(0, 3).Copy
DRowEnd.Offset(0, 7).PasteSpecial xlPasteAll
MbodyNum.Cells(j).Offset(0, 8).Copy
DRowEnd.Offset(0, 6).PasteSpecial xlPasteAll
MbodyNum.Cells(j).Offset(0, 17).Copy
DRowEnd.Offset(0, 8).PasteSpecial xlPasteAll
MbodyNum.Cells(j).Offset(0, 5).Copy
DRowEnd.Offset(0, 9).PasteSpecial xlPasteAll
MbodyNum.Cells(j).Offset(0, 4).Copy
DRowEnd.Offset(0, 10).PasteSpecial xlPasteAll
If MbodyNum.Cells(j).Offset(0, 20).Value = "X" Then
DRowEnd.Offset(0, 11).Value = "売却"
DRowEnd.Interior.Color = RGB(0, 0, 20)
SyaYo.Sheets(I).Activate
Set DRowEnd = DRowEnd.Offset(1, 0)
Else
SyaYo.Sheets(I).Activate
DRowEnd.Value = Application.WorksheetFunction.CountA(DRowEnd, Range(Range("a7"), Range("a7").End(xlDown))) + 1
Set DRowEnd = DRowEnd.Offset(1, 0)
End If

j = j + 1

Else
j = j + 1
End If
End If

Loop

I = I + 1

Loop


SyaYo.Activate

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

ActiveSheet.Sort.SortFields.Clear
   ActiveSheet.Sort.SortFields.Add Key:=Range(Range("F7"), Range("f7").End(xlDown)), _
        SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range(Range("b7"), Range("b7").End(xlDown).Offset(0, 9))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Range(Selection.End(xlDown).Offset(1, 0), Selection.End(xlDown).Offset(1, 10)).Select
Selection.ClearFormats

If ActiveSheet.Name = SheetD.Name Then
Range("d3").Value = Application.WorksheetFunction.Count(Columns(1)) & "台"
Range("a7").Select
Selection.End(xlDown).End(xlDown).Select
Do Until Selection.Value = ""
Range(Selection, Selection.End(xlDown).Offset(0, 10)).Select
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
Selection.End(xlDown).End(xlDown).Select
Loop
Else
Range("d3").Value = Range("a7").End(xlDown).Value & "台"
End If
Range("a7").Activate
Next k

End If
Master.Close False
Application.ScreenUpdating = True
SyaYo.Sheets(1).Activate
ThisWorkbook.SaveCopyAs (strFP & "/autosave/" & Replace(dateT, "/", "") & " ワイズ本社　社用車一覧.xlsm")


End Sub
