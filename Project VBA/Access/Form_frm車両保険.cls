VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm車両保険"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub cmddel_Click()
On Error GoTo cmddel_click_err

    If IsNull(Me.tbBodyNum) Or IsNull(Me.tbInsuNum) Or IsNull(Me.lstInsurance) Then
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
rds.Open "select * from [車両台帳] where [車台番号] = '" & Me.lstDaicho.Column(2) & "'"
If IsNull(lstDaicho.Column(2)) Then
Me.tbIcarnum = ""
Me.tbBodyNum = ""
Me.tbInsuNum = ""
Me.tbstartdate = ""
Me.tbEndDate = ""
Exit Sub
Else
bodynum = rds("車台番号")
Me.tbBodyNum = bodynum
rds.Close

countrecord = DCount("id", "車両保険", "id='" & Me.tbBodyNum & "'")
 
If countrecord > 0 Then

rds.Open "select * from [車両保険] where [ID] = '" & Me.lstDaicho.Column(2) & "'"

Me.tbBodyNum = rds("id")
Me.tbInsuNum = rds("証券番号")
Me.tbstartdate = rds("始期日")
Me.tbEndDate = rds("満期日")
rds.Close

Else
Me.tbInsuNum.SetFocus

End If
End If
Set rds = Nothing
End Sub

Private Sub tbICarnum_AfterUpdate()

Dim sql As String
Dim ds As Database
Dim rs As Recordset
Dim rstC As DAO.Recordset

If IsNull(tbIcarnum) Then
Exit Sub
Else

sql = " select 車両台帳.登録, 車両台帳.番号, 車両台帳.車台番号"
sql = sql & " from 車両台帳"
sql = sql & " Where ((車両台帳.番号) = " & [Forms]![frm車両保険]![tbIcarnum] & ")"
Me.lstDaicho.RowSource = sql


DoCmd.Requery "lstdaicho"

sql = " select count(車両台帳.登録) AS 登録番号の個数"
sql = sql & " From 車両台帳"
sql = sql & " Where ((車両台帳.番号) = " & [Forms]![frm車両保険]![tbIcarnum] & ")"

 
 Set ds = CurrentDb
 Set rs = ds.OpenRecordset(sql)
 
 Me.lblResultRecordCount.Caption = "(" & rs(0) & ")件検索"
 
 tbIcarnum.SetFocus
End If


End Sub


'------------------------------------------------------------
' cmdupdate_Click
'
'------------------------------------------------------------
Private Sub cmdupdate_Click()
Dim rst As ADODB.Recordset
Dim startdate As Date, enddate As Date
On Error GoTo cmdupdate_Click_Err

Set rst = New ADODB.Recordset
Me.Form.AllowAdditions = True
Me.Form.AllowEdits = True
startdate = Me.tbstartdate
enddate = Me.tbEndDate

If Me.lblResultRecordCount.Caption = "(0)件検索" Then
    If Me.tbBodyNum = "" Then
        Exit Sub
    Else
        rst.Open "車両保険", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
        With rst
            .AddNew
            !証券番号 = Me.tbInsuNum
            !始期日 = startdate
            !満期日 = enddate
            .Update
            .Save

            .Requery

            .Close

        End With
        lstInsurance.Requery
    End If
Else

    rst.Open "車両保険", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
    rst.MoveFirst
    Do
        If rst!ID = Me.tbBodyNum Then

            With rst
                !証券番号 = Me.tbInsuNum
                !始期日 = startdate
                !満期日 = enddate
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
            rst.Open "車両保険", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
        With rst
            .AddNew
            !ID = Me.tbBodyNum
            !証券番号 = Me.tbInsuNum
            !始期日 = startdate
            !満期日 = enddate
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


