VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm基本情報"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'------------------------------------------------------------
' cmdEdit_Click
'
'------------------------------------------------------------
Private Sub cmdEdit_Click()
On Error GoTo cmdEdit_Click_Err

    On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If


cmdEdit_Click_Exit:
    Exit Sub

cmdEdit_Click_Err:
    MsgBox Error$
    Resume cmdEdit_Click_Exit

End Sub



'------------------------------------------------------------
' Delete_Click
'
'------------------------------------------------------------
Private Sub Delete_Click()
On Error GoTo Delete_Click_Err

    On Error Resume Next
    DoCmd.GoToControl Screen.PreviousControl.Name
    Err.Clear
    If (Not Form.NewRecord) Then
        DoCmd.RunCommand acCmdDeleteRecord
    End If
    If (Form.NewRecord And Not Form.Dirty) Then
        Beep
    End If
    If (Form.NewRecord And Form.Dirty) Then
        DoCmd.RunCommand acCmdUndo
    End If
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If


Delete_Click_Exit:
    Exit Sub

Delete_Click_Err:
    MsgBox Error$
    Resume Delete_Click_Exit

End Sub


'------------------------------------------------------------
' 作成_Click
'
'------------------------------------------------------------
Private Sub 作成_Click()
On Error GoTo 作成_Click_Err

    On Error Resume Next
    DoCmd.GoToRecord , "", acNewRec
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If


作成_Click_Exit:
    Exit Sub

作成_Click_Err:
    MsgBox Error$
    Resume 作成_Click_Exit

End Sub


'------------------------------------------------------------
' btnPrint_Click
'
'------------------------------------------------------------
Private Sub btnPrint_Click()
On Error GoTo btnPrint_Click_Err

    DoCmd.RunCommand acCmdPrint


btnPrint_Click_Exit:
    Exit Sub

btnPrint_Click_Err:
    MsgBox Error$
    Resume btnPrint_Click_Exit

End Sub


'------------------------------------------------------------
' btnPrev_Click
'
'------------------------------------------------------------
Private Sub btnPrev_Click()
On Error GoTo btnPrev_Click_Err

    On Error Resume Next
    DoCmd.GoToRecord , "", acPrevious
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If


btnPrev_Click_Exit:
    Exit Sub

btnPrev_Click_Err:
    MsgBox Error$
    Resume btnPrev_Click_Exit

End Sub


'------------------------------------------------------------
' btnNext_Click
'
'------------------------------------------------------------
Private Sub btnNext_Click()
On Error GoTo btnNext_Click_Err

    On Error Resume Next
    DoCmd.GoToRecord , "", acNext
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If


btnNext_Click_Exit:
    Exit Sub

btnNext_Click_Err:
    MsgBox Error$
    Resume btnNext_Click_Exit

End Sub


'------------------------------------------------------------
' btngotoLast_Click
'
'------------------------------------------------------------
Private Sub btngotoLast_Click()
On Error GoTo btngotoLast_Click_Err

    DoCmd.GoToRecord , "", acLast


btngotoLast_Click_Exit:
    Exit Sub

btngotoLast_Click_Err:
    MsgBox Error$
    Resume btngotoLast_Click_Exit

End Sub


'------------------------------------------------------------
' btngotoFirst_Click
'
'------------------------------------------------------------
Private Sub btngotoFirst_Click()
On Error GoTo btngotoFirst_Click_Err

    DoCmd.GoToRecord , "", acFirst


btngotoFirst_Click_Exit:
    Exit Sub

btngotoFirst_Click_Err:
    MsgBox Error$
    Resume btngotoFirst_Click_Exit

End Sub

Public Bodynum As Variant

Private Sub cmdDetail_Click()
On Error GoTo cmdDetail_Click_Err
    DoCmd.OpenForm "frm一覧情報", acNormal, , , acEdit, acDialog

cmdDetail_Click_Exit:
    Exit Sub

cmdDetail_Click_Err:
    MsgBox Error$
    Resume cmdDetail_Click_Exit

End Sub


Private Sub cmdExcel_click()
Dim xl As Object
Dim xlWorkbook As Excel.Workbook
Dim xlWorkbook1 As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim strFileName As String
Dim strFileName_2 As String
Dim strPathName As String
Dim TargetRange As Excel.Range
Dim wkbOpen As Boolean
Dim strSQL As String
Dim rds As ADODB.Recordset
Dim Bodynum As String



strPathName = Application.CurrentProject.Path & "\"
strFileName = Dir(strPathName & "*.xlsx")
Do While strFileName <> ""
    If strFileName = "点検表.xlsx" Then
        wkbOpen = True
        Exit Do
    End If
    strFileName = Dir
Loop

If wkbOpen = False Then
    MsgBox "エクセルファイルが見つかれませんでした"
End If

wkbOpen = False
On Error Resume Next

Set xl = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
    Set xl = CreateObject("Excel.Application")    'ｿ｢ｼｿﾀｻ ｽﾇﾇ
Else
    Set xlWorkbook = xl.ActiveWorkbook
    For Each xlWorkbook1 In xl.Workbooks    'ｿｭｷﾁﾁ・ｿ｢ｼｿﾀｻ ｺｳﾇﾏｿｩ ｿﾏｴﾂ ﾆﾄﾀﾏﾀﾌ ｿｭｷﾈｴﾂﾁ・ﾈｮﾀﾎ
        If xlWorkbook1.Name = "点検表.xlsx" Then
            wkbOpen = True
            Exit For
        End If
    Next xlWorkbook1
End If
xl.Visible = True    ' xl.Visible = Falseｴﾂ ﾈｭｸ鯀｡ ｿ｢ｼｿ ｺｸﾀﾌﾁ・ｾﾊｰ・ｳｻｺﾎﾀ釥ｸｷﾎ ﾀﾛｾ
If wkbOpen = False Then    'ｿﾏｴﾂ ﾆﾄﾀﾏﾀﾌ ｿｭｷﾁﾀﾖﾁ・ｾﾊﾀｸｸ
    Set xlWorkbook = xl.Workbooks.Open(strPathName & "点検表.xlsx")
End If

Set rds = New ADODB.Recordset
Set rds.ActiveConnection = CurrentProject.Connection
rds.LockType = adLockPessimistic


rds.Open "select * from [車両台帳] where [車台番号] =  '" & Forms!frm基本情報!tbbodyNum & "'"


If rds("Nox・PM法") = "-" Then
    Set xlSheet = xlWorkbook.Worksheets("被けん引車２")
    xlSheet.Activate

'pgKind
    xlSheet.Range("b5").Value = "3か月点検整備"
    'top
    xlSheet.Range("n3").Value = rds("登録") & rds("番号")
    xlSheet.Range("s3").Value = rds("車名")
    xlSheet.Range("s5").Value = rds("初度登録")
    xlSheet.Range("k3").Value = rds("使用者")

    '2nd

    xlSheet.Range("k5").Value = rds("住所")
    xlSheet.Range("n5").Value = rds("車台番号")
    xlSheet.Range("t2").Value = rds("備考")

    'address
    xlSheet.Range("k42").Value = rds("住所")

    'company
    xlSheet.Range("k45").Value = rds("使用者")
    xlSheet.Range("k48").Value = rds("点検者")

    rds.Close

Else

Set xlSheet = xlWorkbook.Worksheets("トラクタ、トラック")
xlSheet.Activate



    'pgKind
    xlSheet.Range("d4").Value = "3"
    'top
    xlSheet.Range("u2").Value = rds("登録") & rds("番号")
    xlSheet.Range("q2").Value = rds("車名")
    xlSheet.Range("t2").Value = rds("初度登録")
    xlSheet.Range("m2").Value = rds("使用者")

    '2nd

    xlSheet.Range("m4").Value = rds("住所")
    xlSheet.Range("q4").Value = rds("車台番号")
    xlSheet.Range("t4").Value = rds("原動機")
    xlSheet.Range("u4").Value = rds("備考")

    'address
    xlSheet.Range("m66").Value = rds("住所")

    'company
    xlSheet.Range("m68").Value = rds("使用者") & "    " & rds("点検者")

    rds.Close
End If

If Err.Number = 1004 Then
    Set TargetRange = xlSheet.Cells(1, 1)

End If

Set rds = Nothing
End Sub
