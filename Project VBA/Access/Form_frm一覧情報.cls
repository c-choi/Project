VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmàÍóóèÓïÒ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmbLoca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
    cmbLoca.Dropdown
    End If
    
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


