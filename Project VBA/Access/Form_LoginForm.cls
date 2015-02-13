VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdCancel_Click()
DoCmd.Quit
End Sub

Private Sub cmdLogin_Click()
If IsNull(Me.‚”‚‚ID) Or IsEmpty(Me.‚”‚‚ID) Or Me.‚”‚‚ID = "" Then

Beep
MsgBox "ID‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢"
DoCmd.GoToControl "tbid"

Else

If IsNull(Me.tbPWD) Or IsEmpty(Me.tbPWD) Or Me.tbPWD = "" Then

Beep
MsgBox "Password‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢"
DoCmd.GoToControl "txtPWD"

Else

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

rs.activeconnection = CurrentProject.Connection
rs.cursortype = adopendynamic
rs.locktype = adlockoptimistic
rs.Open "select * from [UserList] where [UserID]='" & Me.‚”‚‚ID & "' and [UserPassword]='" & Me.tbPWD & "' "

If rs.BOF Then

Beep
MsgBox "“ü—Í‚µ‚½ID‚ª³‚µ‚­‚ ‚è‚Ü‚¹‚ñ"
Me.‚”‚‚ID = Null
Me.tbPWD = Null
DoCmd.GoToControl "tbID"

Else
Beep
MsgBox rs("UserID") & "‚ÅƒƒOƒCƒ“‚µ‚Ü‚µ‚½B"
DoCmd.Close
DoCmd.OpenForm "Ô—¼ˆê——"
End If

End If

End If

End Sub

Private Sub Loginform_Load()
Const DB_Boolean As Long = 1
ChangeProperty "allowBypasskey", DB_Boolean, False

DoCmd.GoToControl "tdID"

End Sub
