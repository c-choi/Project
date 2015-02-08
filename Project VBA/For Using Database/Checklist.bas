Attribute VB_Name = "Checklist"

Sub CheckList()

Dim Master As Worksheet, CheckList As Worksheet
Dim IDRow As Range
Dim CarInfo As Range, CarDup As Range
Dim PlateNum As String, FirstReg As String, CarName As String, BodyNum As String, Compname As String, Engine As String
Dim Address As String, President As String


Set Master = Worksheets("車両一覧")
Set CheckList = Worksheets("トラクタ、トラック")
Set Target = Range("x2")
Application.ScreenUpdating = False
Master.Activate
On Error Resume Next
Range("a2").Select
Set IDRow = Range(Selection, Selection.End(xlDown))
IDRow.Select
Set CarInfo = IDRow.Find(what:=Target, lookat:=xlWhole, searchdirection:=xlNext)

If CarInfo Is Nothing Then
    MsgBox "Wrong ID number", vbCritical, "Wrong Input"
    CheckList.Activate

Else
    Set CarDup = IDRow.FindNext(CarInfo)


    PlateNum = CarInfo.Offset(0, 3).Value
    FirstReg = CarInfo.Offset(0, 4).Value
    CarName = CarInfo.Offset(0, 6).Value
    BodyNum = CarInfo.Offset(0, 7).Value
    Compname = CarInfo.Offset(0, 11).Value
    Engine = CarInfo.Offset(0, 19).Value
    Address = CarInfo.Offset(0, 28).Value
    President = CarInfo.Offset(0, 29).Value

    If CarDup.Offset(0, 1) = CarInfo.Offset(0, 1) Then
        CheckList.Activate
        Range("m2").Value = Compname
        Range("q2").Value = PlateNum
        Range("t2").Value = CarName
        Range("u2").Value = FirstReg
        Range("m4").Value = Address
        Range("q4").Value = BodyNum
        Range("t4").Value = Engine
        Range("m66").Value = Address
        Range("m66:s66").Select
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
        End With
        Range("m68").Value = Compname & "             " & President

        Range("M68:S71").Select
        Selection.Merge
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = True
        End With
        CheckList.Activate
        Range("x2").Activate

    Else
        Dim DupCount As Integer
        DupCount = Application.CountIf(IDRow, CarInfo)
        Do While i < DupCount
            Dim DupCheck As Integer
            DupCheck = MsgBox(DupCount & " Same ID Number Available" & vbCr & "Click Yes to proceed with " & PlateNum & " Data", vbYesNo, "Duplicate Data")
            Select Case DupCheck
                Case Is = vbYes
                    CheckList.Activate
                    Range("m2").Value = Compname
                    Range("q2").Value = PlateNum
                    Range("t2").Value = CarName
                    Range("u2").Value = FirstReg
                    Range("m4").Value = Address
                    Range("q4").Value = BodyNum
                    Range("t4").Value = Engine
                    Range("m66").Value = Address
                    Range("m66:s66").Select
                    With Selection
                        .HorizontalAlignment = xlLeft
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = True
                    End With
                    Range("m68").Value = Compname & "             " & President

                    Range("M68:S71").Select
                    Selection.Merge
                    With Selection
                        .HorizontalAlignment = xlLeft
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = True
                    End With


                    CheckList.Activate
                    Range("x2").Activate
                    Exit Sub

                Case Is = vbNo
                    Set CarDup = IDRow.FindNext(CarInfo)
                    Set CarInfo = CarDup


                    PlateNum = CarInfo.Offset(0, 3).Value
                    FirstReg = CarInfo.Offset(0, 4).Value
                    CarName = CarInfo.Offset(0, 6).Value
                    BodyNum = CarInfo.Offset(0, 7).Value
                    Compname = CarInfo.Offset(0, 11).Value
                    Engine = CarInfo.Offset(0, 19).Value
                    Address = CarInfo.Offset(0, 28).Value
                    President = CarInfo.Offset(0, 29).Value

                    CheckList.Activate
                    Range("m2").Value = Compname
                    Range("q2").Value = PlateNum
                    Range("t2").Value = CarName
                    Range("u2").Value = FirstReg
                    Range("m4").Value = Address
                    Range("q4").Value = BodyNum
                    Range("t4").Value = Engine
                    Range("m66").Value = Address
                    Range("m66:s66").Select
                    With Selection
                        .HorizontalAlignment = xlLeft
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = True
                    End With
                    Range("m68").Value = Compname & "             " & President

                    Range("M68:S71").Select
                    Selection.Merge
                    With Selection
                        .HorizontalAlignment = xlLeft
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                        .MergeCells = True
                    End With
            End Select
            i = i + 1
        Loop
    End If
End If

CheckList.Activate
Range("x2").Activate
Application.ScreenUpdating = True

End Sub


Sub DaiSyaCheckList()
``` Only Copied version (not fixed)

Dim Master As Worksheet, CheckList As Worksheet
Dim IDRow As Range
Dim CarInfo As Range
Dim PlateNum As String, FirstReg As String, CarName As String, BodyNum As String, Compname As String, Engine As String
Dim Address As String, President As String


Set Master = Worksheets("車両一覧")
Set CheckList = Worksheets("被牽引車")
Set Target = Range("z2")
Application.ScreenUpdating = False
Master.Activate
On Error Resume Next
Range("a2").Select
Set IDRow = Range(Selection, Selection.End(xlDown))
IDRow.Select
Set CarInfo = IDRow.Find(what:=Target, lookat:=xlWhole)

If CarInfo Is Nothing Then
MsgBox "Wrong ID number", vbCritical, "Wrong Input"
CheckList.Activate

Else
PlateNum = CarInfo.Offset(0, 3).Value
FirstReg = CarInfo.Offset(0, 4).Value
CarName = CarInfo.Offset(0, 6).Value
BodyNum = CarInfo.Offset(0, 7).Value
Compname = CarInfo.Offset(0, 11).Value
Engine = CarInfo.Offset(0, 19).Value
Address = CarInfo.Offset(0, 28).Value
President = CarInfo.Offset(0, 29).Value

CheckList.Activate
Range("n2").Value = CarInfo.Value
Range("r2").Value = PlateNum
Range("v2").Value = CarName
Range("w2").Value = FirstReg
Range("n4").Value = Address
Range("r4").Value = BodyNum
Range("v4").Value = Engine
Range("m66").Value = Address
    Range("m66:s66").Select
With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
Range("m68").Value = Compname & Chr(10) & "" & Chr(13) & President

    Range("M68:S71").Select
        Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
End If
Application.ScreenUpdating = True

End Sub




