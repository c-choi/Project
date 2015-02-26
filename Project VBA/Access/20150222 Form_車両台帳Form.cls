VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Ô—¼‘ä’ Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public BodyNum As Variant
 
Private Sub cmdDetail_Click()
BodyNum = tbBodyNum.Value
Call DoCmd.OpenForm("frmˆê——î•ñ", acNormal, , , acFormEdit, acDialog)

Forms!frmˆê——î•ñ!tbBodyNum.Value = BodyNum

End Sub
