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
Public CodeNum As Variant


Private Sub cmddel_Click()
On Error GoTo cmddel_click_err

    If IsNull(Me.tbbodyNum) Or IsNull(Me.tbInsuNum) Or IsNull(Me.lstInsurance) Then
    MsgBox "削除する情報を選べてください。"
    Exit Sub
    End If
    
    Dim sql, msg, style, title, help, ctxt, response, message, default, myvalue
    
    msg = Me.tbbodyNum & "の保険情報を削除しますか？"
    style = vbYesNo + vbCritical + vbDefaultButton1
    title = "削除"
    
    response = MsgBox(msg, style, title)
    
    If response = vbYes Then
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = CurrentProject.Connection
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockOptimistic
    
    rs.Open "select * from [車両保険] where [id] = '" & Me.tbbodyNum & "'"
    
    Do Until rs.EOF
     
     rs.Delete
     rs.MoveNext
    Loop
    
    Set rs = Nothing
    
    MsgBox Me.tbbodyNum & "の保険情報を削除しました。"
    DoCmd.Requery "lstInsurance"
    
Else
 MsgBox "閉じます。"
    End If
    
cmddel_click_err:
    If Err.Number <> 0 Then
    MsgBox "Error" & Err.Number & " " & Err.Description
    End If
End Sub


Private Sub lstInsurance_BeforeUpdate(Cancel As Integer)

Dim sql As String
Dim ds As Database
Dim rs As Recordset
Dim rstC As DAO.Recordset

sql = " select count(*) AS 保険の個数"
sql = sql & " From 車両保険;"

  
  Set ds = CurrentDb
 Set rs = ds.OpenRecordset(sql)
 Me.lbllstrecordcount.Caption = "(" & rs(0) & ")件検索"
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
Me.tbComment = ""
rds.Open "select * from [車両台帳] where [車台番号] = '" & Me.lstDaicho.Column(2) & "'"
If IsNull(lstDaicho.Column(2)) Then
Me.tbICarnum = ""
Me.tbbodyNum = ""
Me.tbInsuNum = ""
Me.tbstartdate = ""
Me.tbEndDate = ""
Me.tbComment = ""
Exit Sub
Else
bodynum = rds("車台番号")
Me.tbbodyNum = bodynum
rds.Close

countrecord = DCount("id", "車両保険", "id='" & Me.tbbodyNum & "'")
 
If countrecord > 0 Then

rds.Open "select * from [車両保険] where [ID] = '" & Me.lstDaicho.Column(2) & "'"

Me.tbbodyNum = rds("id")
Me.tbInsuNum = rds("証券番号")
Me.tbstartdate = rds("始期日")
Me.tbEndDate = rds("満期日")
Me.tbComment = rds("備考")
rds.Close

Else
Me.tbInsuNum.SetFocus

End If
End If
Set rds = Nothing
End Sub


Private Sub lstDaicho_exit(Cancel As Integer)
Dim insunum As String

On Error Resume Next
insunum = Me.tbInsuNum
If insunum = "" Then

If IsEmpty(CodeNum) Then
Me.tbInsuNum = insunum
Exit Sub
Else
Me.tbInsuNum = CodeNum
End If
Else
Exit Sub
End If
End Sub


Private Sub tbICarnum_AfterUpdate()

Dim sql As String
Dim ds As Database
Dim rs As Recordset
Dim rstC As DAO.Recordset

If IsNull(tbICarnum) Then
Exit Sub
Else

sql = " select 車両台帳.登録, 車両台帳.番号, 車両台帳.車台番号"
sql = sql & " from 車両台帳"
sql = sql & " Where ((車両台帳.番号) = " & [Forms]![frm車両保険]![tbICarnum] & ")"
Me.lstDaicho.RowSource = sql


DoCmd.Requery "lstdaicho"

sql = " select count(車両台帳.登録) AS 登録番号の個数"
sql = sql & " From 車両台帳"
sql = sql & " Where ((車両台帳.番号) = " & [Forms]![frm車両保険]![tbICarnum] & ")"

 
 Set ds = CurrentDb
 Set rs = ds.OpenRecordset(sql)
 
 Me.lblResultRecordCount.Caption = "(" & rs(0) & ")件検索"
 
 tbICarnum.SetFocus
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

CodeNum = Left(Me.tbInsuNum, InStr(Me.tbInsuNum, "-"))

If Me.lblResultRecordCount.Caption = "(0)件検索" Then
    If Me.tbbodyNum = "" Then
        Exit Sub
    Else
        rst.Open "車両保険", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
        With rst
            .AddNew
            !証券番号 = Me.tbInsuNum
            !始期日 = startdate
            !満期日 = enddate
            !備考 = Me.tbComment
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
        If rst!ID = Me.tbbodyNum Then

            With rst
                !証券番号 = Me.tbInsuNum
                !始期日 = startdate
                !満期日 = enddate
                !備考 = Me.tbComment
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
            !ID = Me.tbbodyNum
            !証券番号 = Me.tbInsuNum
            !始期日 = startdate
            !満期日 = enddate
            !備考 = Me.tbComment
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


Private Sub tbstartdate_AfterUpdate()
Dim StartD As Date
Dim endD As Date
On Error Resume Next
StartD = Me.tbstartdate.Value
endD = CLng(StartD) + 1 * 365
If Day(endD) <> Day(StartD) Then
endD = endD + 1
End If
Me.tbEndDate = endD
If Err.Number <> 0 Then
Exit Sub
Else
End If
End Sub

Private Sub tbstartdate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
    tbstartdate.Dropdown
    End If
    
End Sub



Private Sub tbenddate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
    tbEndDate.Dropdown
    End If
    
End Sub



Private Sub tbinsunum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
    tbInsuNum.Dropdown
    End If
    
End Sub


