VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_é–ópé‘Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDetail_Click()
On Error GoTo cmdDetail_Click_Err
    DoCmd.OpenForm "frmàÍóóèÓïÒé–ópé‘", acNormal, , , acEdit, acDialog

cmdDetail_Click_Exit:
    Exit Sub

cmdDetail_Click_Err:
    MsgBox Error$
    Resume cmdDetail_Click_Exit

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


