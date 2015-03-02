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
Dim bodynum As Variant

Set rds = New ADODB.Recordset
Set rds.ActiveConnection = CurrentProject.Connection
rds.LockType = adLockPessimistic
Me.tbInsuNum = ""
Me.tbstartdate = ""
Me.tbEndDate = ""
rds.Open "select * from [Ô—¼‘ä’ ] where [Ô‘ä”Ô†] = '" & Me.lstDaicho.Column(2) & "'"
If IsNull(lstDaicho.Column(2)) Then
Me.tbIcarnum = ""
Me.tbBodyNum = ""
Me.tbInsuNum = ""
Me.tbstartdate = ""
Me.tbEndDate = ""
Exit Sub
Else
bodynum = rds("Ô‘ä”Ô†")
Me.tbBodyNum = bodynum
rds.Close

countrecord = DCount("id", "Ô—¼•ÛŒ¯", "id='" & Me.tbBodyNum & "'")
 
If countrecord > 0 Then

rds.Open "select * from [Ô—¼•ÛŒ¯] where [ID] = '" & Me.lstDaicho.Column(2) & "'"

Me.tbBodyNum = rds("id")
Me.tbInsuNum = rds("ØŒ””Ô†")
Me.tbstartdate = rds("nŠú“ú")
Me.tbEndDate = rds("–Šú“ú")
rds.Close

Else
Me.tbInsuNum.SetFocus

End If
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
Me.Form.AllowAdditions = True
Me.Form.AllowEdits = True
If Me.lblResultRecordCount.Caption = "(0)ŒŒŸõ" Then
    If Me.tbBodyNum = "" Then
        Exit Sub
    Else
        rst.Open "Ô—¼•ÛŒ¯", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
        With rst
            .AddNew
            !ØŒ””Ô† = Me.tbInsuNum
            !nŠú“ú = Me.tbstartdate
            !–Šú“ú = Me.tbEndDate
            .Update
            .Save

            .Requery

            .Close

        End With
        lstInsurance.Requery
    End If
Else

    rst.Open "Ô—¼•ÛŒ¯", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
    rst.MoveFirst
    Do
        If rst!ID = Me.tbBodyNum Then

            With rst
                !ØŒ””Ô† = Me.tbInsuNum
                !nŠú“ú = Me.tbstartdate
                !–Šú“ú = Me.tbEndDate
                .Update

                .Requery
                .Close
                lstInsurance.Requery
            End With
            Set rst = Nothing
            Exit Sub

        End If
        rst.MoveNext
    Loop Until rst.EOF
            
          If rst.EOF Then
          rst.Close
            rst.Open "Ô—¼•ÛŒ¯", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
        With rst
            .AddNew
            !ID = Me.tbBodyNum
            !ØŒ””Ô† = Me.tbInsuNum
            !nŠú“ú = Me.tbstartdate
            !–Šú“ú = Me.tbEndDate
            .Update
            .Save

            .Requery

            .Close

        End With
            End If
        lstInsurance.Requery
        
    Set rst = Nothing

End If



cmdupdate_Click_Exit:

Exit Sub

cmdupdate_Click_Err:
MsgBox Error$
Resume cmdupdate_Click_Exit

End Sub




cmdDelete_Click_Err:
    MsgBox Error$
    Resume cmdDelete_Click_Exit

End Sub


