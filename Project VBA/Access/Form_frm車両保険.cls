VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmÔ—¼•ÛŒ¯"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub lstDaicho_DblClick(Cancel As Integer)
Dim rds As ADODB.Recordset
Dim countrecord As Variant
Set rds = New ADODB.Recordset
Set rds.ActiveConnection = CurrentProject.Connection
rds.LockType = adLockPessimistic

rds.Open "select * from [Ô—¼‘ä’ ] where [Ô‘ä”Ô†] = '" & Me.lstDaicho.Column(2) & "'"
Me.tbBodyNum = rds("Ô‘ä”Ô†")

rds.Close

countrecord = DCount("id", "Ô—¼•ÛŒ¯", "id='" & Me.tbBodyNum & "'")
 
If countrecord > 0 Then

rds.Open "select * from [Ô—¼•ÛŒ¯] where [ID] = '" & Me.lstDaicho.Column(2) & "'"

Me.tbBodyNum = rds("id")
Me.tbInsuNum = rds("ØŒ””Ô†")
Me.tbstartdate = rds("nŠú“ú")
Me.tbEndDate = rds("–Šú“ú")

Else
rds.Open "select * from [Ô—¼•ÛŒ¯] where [ID] = '" & Me.lstDaicho.Column(2) & "'"
rds.AddNew
Me.tbInsuNum = ""
Me.tbstartdate = ""
Me.tbEndDate = ""
Me.BxReg = Me.lstDaicho.Column(0)
Me.bxNum = Me.lstDaicho.Column(1)
Me.bxbodynum = Me.tbBodyNum

Me.tbInsuNum.SetFocus

End If
Set rds = Nothing
End Sub

Private Sub tbICarnum_AfterUpdate()

Dim SQL As String
Dim ds As Database
Dim rs As Recordset
Dim rstC As DAO.Recordset

If IsNull(tbIcarnum) Then
Exit Sub
Else

SQL = " select Ô—¼‘ä’ .“o˜^, Ô—¼‘ä’ .”Ô†, Ô—¼‘ä’ .Ô‘ä”Ô†"
SQL = SQL & " from Ô—¼‘ä’ "
SQL = SQL & " Where ((Ô—¼‘ä’ .”Ô†) = " & [Forms]![frmÔ—¼•ÛŒ¯]![tbIcarnum] & ")"
Me.lstDaicho.RowSource = SQL


DoCmd.Requery "lstdaicho"

SQL = " select count(Ô—¼‘ä’ .“o˜^) AS “o˜^”Ô†‚ÌŒÂ”"
SQL = SQL & " From Ô—¼‘ä’ "
SQL = SQL & " Where ((Ô—¼‘ä’ .”Ô†) = " & [Forms]![frmÔ—¼•ÛŒ¯]![tbIcarnum] & ")"

 
 Set ds = CurrentDb
 Set rs = ds.OpenRecordset(SQL)
 
 Me.lblResultRecordCount.Caption = "(" & rs(0) & ")ŒŒŸõ"
 
 tbIcarnum.SetFocus
End If


End Sub


'------------------------------------------------------------
' cmdupdate_Click
'
'------------------------------------------------------------
Private Sub cmdupdate_Click()
Dim rst As ADODB.Recordset
On Error GoTo cmdupdate_Click_Err
Set rst = New ADODB.Recordset

    rst.Open "Ô—¼•ÛŒ¯", CurrentProject.Connection, adOpenDynamic, adLockOptimistic

    With rst
    .Save
    !ID = Me.bxbodynum
    !ØŒ””Ô† = Me.tbInsuNum
    !nŠú“ú = Me.tbstartdate
    !–Šú“ú = Me.tbEndDate
    
    .Update
    rst.Requery
    .Close
    End With
    
    Set rst = Nothing

cmdupdate_Click_Exit:
    Exit Sub

cmdupdate_Click_Err:
    MsgBox Error$
    Resume cmdupdate_Click_Exit

End Sub



