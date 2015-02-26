VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmŽÔ—¼“ü—Í"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmbLoca_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
    cmbLoca.Dropdown
    End If
    
End Sub

Private Sub tbPlateNum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub

Private Sub tbType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If

End Sub


