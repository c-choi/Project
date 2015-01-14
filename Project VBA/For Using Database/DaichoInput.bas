Attribute VB_Name = "DaichoInput"
Sub DaichoInput()

Dim Daicho As Workbook
Dim DbodyC As Integer
Dim DRowEnd As Range, DRowEndD As Range, DrowEndS As Range
Dim MbodyNum As Range, CarRng As Range
Dim SheetC As Integer
Dim sheetN As String
Dim Master As Workbook
Dim dateT As Date
Dim SheetD As Worksheet
Dim CountDumpW As Range, CountDumpS As Range, countdumpF As Range
Dim strAddW As String, strAddS As String, strAddF As String
Dim DumpRng As Range
Dim strFP As String
Dim CarCount As Integer

strFP = ThisWorkbook.Path
dateT = Format(Now, YYYY - MM - DD)

Set Daicho = ThisWorkbook
Application.DisplayAlerts = False
Application.ScreenUpdating = False
On Error Resume Next
Set Master = Workbooks("���C�Y�E�Z�u���}�X�^�t�@�C��.xlsm")
If Err.Number <> 0 Then
    Filename = Application.GetOpenFilename
    If Filename = False Then
        Exit Sub
    Else
        Workbooks.Open Filename:=Filename, ReadOnly:=True
        DaichoInput
    End If
Else
    Master.Sheets(1).Activate
    Set MbodyNum = Range(Range("h2"), Range("h2").End(xlDown))
    Daicho.Activate
    I = 1
    j = 1

    SheetC = Sheets.Count
    Set SheetD = Daicho.Sheets("�_���v�ۗL�ꗗ")
    SheetD.Activate
    SheetD.Range("a7").Activate
    Set DRowEndD = ActiveCell
    Range(DRowEndD, DRowEndD.End(xlDown).End(xlDown).End(xlDown).Offset(0, 11)).Select
    Selection.Clear

    Master.Sheets(1).Activate
    Z = 0
    Set CountDumpW = MbodyNum.Offset(0, 11).Find(what:="���C�Y�_���v", lookat:=xlPart)

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
    Selection.Value = "�Z�u���@�ۗL�ԗ�"
    Selection.Offset(1, -1).Select
    Set DrowEndS = Selection
    DrowEndS.PasteSpecial xlPasteAll

    y = 0
    Set CountDumpS = MbodyNum.Offset(0, 11).Find(what:="�Z�u���_���v", lookat:=xlPart)

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
    Selection.Value = "�z�C-���N��-��"
    Selection.Offset(1, -1).Select
    Selection.PasteSpecial xlPasteAll

    f = 0
    Set countdumpF = MbodyNum.Offset(0, 11).Find(what:="�z�C-���N��-��", lookat:=xlPart)

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




    Do Until I = SheetC
        Daicho.Activate
        Sheets(I).Activate
        sheetN = ActiveSheet.Name
        Range("a7").Activate
        Set DRowEnd = ActiveCell
        Range(DRowEnd, DRowEnd.End(xlDown).End(xlDown).End(xlDown).Offset(0, 11)).Select
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
                If MbodyNum.Cells(j).Offset(0, 20).Value = "X" Then
                    DRowEnd.Offset(0, 11).Value = "���p"
                    DRowEnd.Interior.Color = rgbDarkGray
                    DRowEnd.Value = "-"
                    Daicho.Sheets(I).Activate
                    Set DRowEnd = DRowEnd.Offset(1, 0)
                Else
                    Daicho.Sheets(I).Activate
                    DRowEnd.Value = Application.WorksheetFunction.CountA(DRowEnd, Range(Range("a7"), Range("a7").End(xlDown))) + 1
                    Set DRowEnd = DRowEnd.Offset(1, 0)
                End If


            ElseIf InStr(MbodyNum.Cells(j).Offset(0, 11).Value, sheetN) Then
                If InStr(MbodyNum.Cells(j).Offset(0, 11).Value, "�_���v") Or _
                   MbodyNum.Cells(j).Offset(0, 11).Value = "�z�C-���N��-��" Then
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
                    If MbodyNum.Cells(j).Offset(0, 20).Value = "X" Then
                        DRowEnd.Offset(0, 11).Value = "���p"
                        DRowEnd.Interior.Color = rgbDarkGray
                        DRowEnd.Value = "-"
                        DRowEndD.Offset(0, 11).Value = "���p"
                        DRowEndD.Interior.Color = rgbDarkGray
                        DRowEndD.Value = "-"
                        Daicho.Sheets(I).Activate
                        Set DRowEnd = DRowEnd.Offset(1, 0)
                        Set DRowEndD = DRowEndD.Offset(1, 0)
                    Else
                        Daicho.Sheets(I).Activate
                        DRowEnd.Value = Application.WorksheetFunction.CountA(DRowEnd, Range(Range("a7"), Range("a7").End(xlDown))) + 1
                        Set DRowEnd = DRowEnd.Offset(1, 0)
                        SheetD.Activate
                        DumpRng.Offset(0, -5).Value = Application.WorksheetFunction.CountA(Range(DumpRng.End(xlUp).Offset(1, 0), DumpRng))
                        Set DRowEndD = DRowEndD.Offset(1, 0)
                    End If

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
                    If MbodyNum.Cells(j).Offset(0, 20).Value = "X" Then
                        DRowEnd.Offset(0, 11).Value = "���p"
                        DRowEnd.Interior.Color = rgbDarkGray
                        DRowEnd.Value = "-"
                        Daicho.Sheets(I).Activate
                        Set DRowEnd = DRowEnd.Offset(1, 0)
                    Else
                        Daicho.Sheets(I).Activate
                        DRowEnd.Value = Application.WorksheetFunction.CountA(DRowEnd, Range(Range("a7"), Range("a7").End(xlDown))) + 1
                        Set DRowEnd = DRowEnd.Offset(1, 0)
                    End If
                End If

            ElseIf InStr(MbodyNum.Cells(j).Offset(0, 11).Value, "�_���v") Or _
                   MbodyNum.Cells(j).Offset(0, 11).Value = "�z�C-���N��-��" Then
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
                If MbodyNum.Cells(j).Offset(0, 20).Value = "X" Then
                    DRowEndD.Offset(0, 11).Value = "���p"
                    DRowEndD.Interior.Color = rgbDarkGray
                    DRowEndD.Value = "-"
                    Set DRowEndD = DRowEndD.Offset(1, 0)
                Else
                    SheetD.Activate
                    DumpRng.Offset(0, -5).Value = Application.WorksheetFunction.CountA(Range(DumpRng.End(xlUp).Offset(1, 0), DumpRng))
                    Set DRowEndD = DRowEndD.Offset(1, 0)
                End If

            ElseIf InStr(MbodyNum.Cells(j).Offset(0, 11).Value, "�_���v") Or _
                   MbodyNum.Cells(j).Offset(0, 11).Value = "�z�C-���N��-��" Then
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
                If MbodyNum.Cells(j).Offset(0, 20).Value = "X" Then
                    DRowEndD.Offset(0, 11).Value = "���p"
                    DRowEndD.Interior.Color = rgbDarkGray
                    DRowEndD.Value = "-"
                    Set DRowEndD = DRowEndD.Offset(1, 0)
                Else
                    SheetD.Activate
                    DumpRng.Offset(0, -5).Value = Application.WorksheetFunction.CountA(Range(DumpRng.End(xlUp).Offset(1, 0), DumpRng))
                    Set DRowEndD = DRowEndD.Offset(1, 0)
                End If
            End If


            j = j + 1

        Loop

        I = I + 1

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

        ActiveSheet.Sort.SortFields.Clear
        ActiveSheet.Sort.SortFields.Add Key:=Range(Range("F7"), Range("f7").End(xlDown)), _
                                        SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:=xlSortNormal
        With ActiveSheet.Sort
            .SetRange Range(Range("a7"), Range("a7").End(xlDown).Offset(0, 11))
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        Range(Selection.End(xlDown).Offset(1, 0), Selection.End(xlDown).Offset(1, 10)).Select
        Selection.ClearFormats

        If ActiveSheet.Name = SheetD.Name Then
            Range("d3").Value = Application.WorksheetFunction.Count(Columns(1)) & "��"
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
             CarCount = Application.WorksheetFunction.Count(Columns(1))
             Range("d3").Value = CarCount & "��"
             Range("a7").Select
             Set CarRng = Range(Selection, Selection.Offset(CarCount - 1, 0))
             Selection.AutoFill Destination:=CarRng, Type:=xlFillSeries
        End If
        Range("a7").Activate
    Next k

End If
Master.Close False
Application.ScreenUpdating = True
Daicho.Sheets(1).Activate
ThisWorkbook.SaveCopyAs (strFP & "/autosave/" & Replace(dateT, "/", "") & " �ԗ��䒠 �S��.xlsm")


End Sub
