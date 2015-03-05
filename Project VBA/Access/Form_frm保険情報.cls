VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm保険情報"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------
' Delete_Click
'
'------------------------------------------------------------
Private Sub cmddel_Click()
On Error GoTo cmddel_click_err

    If IsNull(Me.tbBodyNum) Or IsNull(Me.tbInsuNum) Then
    MsgBox "削除する情報を選べてください。"
    Exit Sub
    End If
    
    Dim sql, msg, style, title, help, ctxt, response, message, default, myvalue
    
    msg = Me.tbBodyNum & "の保険情報を削除しますか？"
    style = vbYesNo + vbCritical + vbDefaultButton1
    title = "削除"
    
    response = MsgBox(msg, style, title)
    
    If response = vbYes Then
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = CurrentProject.Connection
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    
    rs.Open "select * from [車両保険] where [id] = '" & Me.tbBodyNum & "'"
    
    Do Until rs.EOF
     
     rs.Delete
     rs.MoveNext
    Loop
    
    Set rs = Nothing
    
    MsgBox Me.tbBodyNum & "の保険情報を削除しました。"
    DoCmd.Requery "lstInsurance"
    
Else
 MsgBox "閉じます。"
    End If
    
cmddel_click_err:
    If Err.Number <> 0 Then
    MsgBox "Error" & Err.Number & " " & Err.Description
    End If
    
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

Public bodynum As Variant

Private Sub cmdDetail_Click()
On Error GoTo cmdDetail_Click_Err
bodynum = [Forms]![frm基本情報]![tbBodyNum]
    DoCmd.OpenForm "frm一覧情報", acNormal, , , acEdit, acDialog

cmdDetail_Click_Exit:
    Exit Sub

cmdDetail_Click_Err:
    MsgBox Error$
    Resume cmdDetail_Click_Exit

End Sub


