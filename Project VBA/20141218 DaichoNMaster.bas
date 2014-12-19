Attribute VB_Name = "DaichoNMaster"
Sub CarBodyNumCompare()
Dim RngOld As Range, RngNew As Range
Dim BodyNum As Range
Dim BodyCount As Integer
Dim Master As Worksheet
Dim DaiCho As Workbook
Dim DaiChoSheet As Worksheet
Dim DaiChoCount As Integer
Dim CarNum As Integer
Dim RngDaicho As Range
Dim i As Integer
Dim k As Integer
Dim j As Integer
Set Master = ThisWorkbook.Sheets(1)

Set DaiCho = Workbooks("車両台帳　全体.xlsx")
DaiChoCount = DaiCho.Sheets.Count
i = 1
k = 1
j = 1
Application.ScreenUpdating = False
Master.Activate
Set RngOld = Master.Range(Range("e2"), Range("e2").End(xlDown))
Set BodyNum = Master.Range(Range("j2"), Range("j2").End(xlDown))
BodyCount = BodyNum.Count
Set RngNew = Master.Range(Range("a2"), Range("a2").Offset(BodyCount, 0))

DaiCho.Activate
Set DaiChoSheet = DaiCho.Sheets(k)
DaiChoSheet.Activate

If DaiChoSheet.Range("f8").Value <> "" Then
Set RngDaicho = DaiChoSheet.Range(Range("f7"), Range("f7").End(xlDown))
CarNum = RngDaicho.Count

Else
Set RngDaicho = DaiChoSheet.Range("f7")

End If

Master.Activate

Do While BodyNum.Cells(j).Value <> ""

    Do Until k = DaiChoCount
        
        Do Until i = CarNum

            If BodyNum.Cells(j).Value = RngDaicho.Cells(i).Value Then
            RngNew.Cells(j).Value = RngDaicho.Cells(i).Offset(0, -4).Value
            i = 1
            DaiCho.Activate
            k = 1
            Set DaiChoSheet = DaiCho.Sheets(k)

                If DaiChoSheet.Range("f8").Value <> "" Then
                DaiChoSheet.Activate
                Set RngDaicho = DaiChoSheet.Range(Range("f7"), Range("f7").End(xlDown))
                CarNum = RngDaicho.Count

                Else
                Set RngDaicho = DaiChoSheet.Range("f7")
                End If
            j = j + 1

            Else
            i = i + 1
            End If

            Loop
        i = 1
        k = k + 1
        DaiCho.Activate
        Set DaiChoSheet = DaiCho.Sheets(k)
        DaiChoSheet.Activate

            If DaiChoSheet.Range("f8").Value <> "" Then
            Set RngDaicho = DaiChoSheet.Range(Range("f7"), Range("f7").End(xlDown))
            CarNum = RngDaicho.Count

            Else
            Set RngDaicho = DaiChoSheet.Range("f7")
            End If

        Loop
        k = 1
        j = j + 1
Loop

Application.ScreenUpdating = True
Master.Activate
End Sub

Sub ComparePlateNum()

Dim OldNum As Range
Dim NewNum As Range
Dim RngOld As Range
Dim RowCount As Integer
Dim i As Integer
Set OldNum = Range(Range("e2"), Range("e2").End(xlDown))
RowCount = OldNum.Offset(0, -1).Count
Set NewNum = OldNum.Offset(0, -4)
Set RngOld = OldNum.Offset(0, 13)
i = 1

Do Until i = RowCount
    If OldNum.Cells(i).Value = NewNum.Cells(i).Value Then
    RngOld.Cells(i).Value = "番号変更X"
    i = i + 1
    Else
        If NewNum.Cells(i).Value = "" Then
        RngOld.Cells(i).Value = "車両台帳データX"
        i = i + 1

        Else
        RngOld.Cells(i).Value = OldNum.Cells(i).Value
        RngOld.Cells(i).Offset(0, 1).Value = NewNum.Cells(i).Value
        i = i + 1
        End If
    End If
Loop
End Sub

