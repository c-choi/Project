VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_車両台帳Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public BodyNum As Variant
 
Private Sub cmdDetail_Click()
BodyNum = tbBodyNum.Value
Call DoCmd.OpenForm("frm一覧情報", acNormal, , , acFormEdit, acDialog)

Forms!frm一覧情報!tbBodyNum.Value = BodyNum

End Sub
