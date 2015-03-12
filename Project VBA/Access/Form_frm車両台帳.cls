VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm車両台帳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmbLoca_AfterUpdate()

Dim sql As String

sql = " select 車両台帳.登録,　車両台帳.番号,　車両台帳.初度登録,　車両台帳.型式,　車両台帳.車名,　車両台帳.車台番号,"
sql = sql & " 車両台帳.種別,　車両台帳.最大積載量,　車両台帳.車両総重量,　車両台帳.Nox・PM法,　車両台帳.基準緩和事項,　車両台帳.場所, 車両台帳.社用車, 車両台帳.ダンプ"
sql = sql & " from 車両台帳"
sql = sql & " Where (((車両台帳.場所) = " & [Forms]![frm車両台帳]![cmbLoca] & ")"
 sql = sql & " and ((車両台帳.社用車) = " & [Forms]![frm車両台帳]![chkCompCar] & ")"
 sql = sql & " and ((車両台帳.ダンプ) = " & [Forms]![frm車両台帳]![chkDump] & "))"
sql = sql & " order by 車両台帳.番号;"
Me.lstDaicho.RowSource = sql

DoCmd.Requery "lstDaicho"

sql = " select count(車両台帳.登録) AS 登録番号の個数"
sql = sql & " From 車両台帳"
sql = sql & " Where (((車両台帳.場所) = " & [Forms]![frm車両台帳]![cmbLoca] & ")"
 sql = sql & " and ((車両台帳.社用車) = " & [Forms]![frm車両台帳]![chkCompCar] & ")"
 sql = sql & " and ((車両台帳.ダンプ) = " & [Forms]![frm車両台帳]![chkDump] & "))"
 
 
 Dim ds As Database
 Dim rs As Recordset
 
 Set ds = CurrentDb
 Set rs = ds.OpenRecordset(sql)
 
 Me.lblResultRecordCount.Caption = "(" & rs(0) & ")件検索"

Me.tbDaichoCarNum = ""

If Me.cmbLoca.Value = "12" Then
chkDump = 0
chkCompCar = -1
chkCompCar_afterupdate
ElseIf Me.cmbLoca.Value = "13" Then
chkDump = 0
chkCompCar = -1
chkCompCar_afterupdate
ElseIf Me.cmbLoca.Value = "10" Then
chkCompCar = 0
chkDump = -1
chkDump_afterupdate
ElseIf Me.cmbLoca.Value = "11" Then
chkCompCar = -1
chkDump = -1
chkDump_afterupdate
ElseIf Me.cmbLoca.Value = "14" Then
chkCompCar = 0
chkDump = -1
chkDump_afterupdate
End If
End Sub


Private Sub chkCompCar_afterupdate()

Dim sql As String

sql = " select 車両台帳.登録,　車両台帳.番号,　車両台帳.初度登録,　車両台帳.型式,　車両台帳.車名,　車両台帳.車台番号,"
sql = sql & " 車両台帳.種別,　車両台帳.最大積載量,　車両台帳.車両総重量,　車両台帳.Nox・PM法,　車両台帳.基準緩和事項,　車両台帳.場所, 車両台帳.社用車, 車両台帳.ダンプ"
sql = sql & " from 車両台帳"
sql = sql & " Where (((車両台帳.場所) = " & [Forms]![frm車両台帳]![cmbLoca] & ")"
 sql = sql & " and ((車両台帳.社用車) = " & [Forms]![frm車両台帳]![chkCompCar] & ")"
 sql = sql & " and ((車両台帳.ダンプ) = " & [Forms]![frm車両台帳]![chkDump] & "))"
sql = sql & " order by 車両台帳.番号;"
Me.lstDaicho.RowSource = sql

DoCmd.Requery "lstDaicho"

sql = " select count(車両台帳.登録) AS 登録番号の個数"
sql = sql & " From 車両台帳"
sql = sql & " Where (((車両台帳.場所) = " & [Forms]![frm車両台帳]![cmbLoca] & ")"
 sql = sql & " and ((車両台帳.社用車) = " & [Forms]![frm車両台帳]![chkCompCar] & ")"
 sql = sql & " and ((車両台帳.ダンプ) = " & [Forms]![frm車両台帳]![chkDump] & "))"
 
 
 Dim ds As Database
 Dim rs As Recordset
 
 Set ds = CurrentDb
 Set rs = ds.OpenRecordset(sql)
 
 Me.lblResultRecordCount.Caption = "(" & rs(0) & ")件検索"

Me.tbDaichoCarNum = ""
Me.chkDump = "0"
End Sub


Private Sub chkDump_afterupdate()

Dim sql As String

sql = " select 車両台帳.登録,　車両台帳.番号,　車両台帳.初度登録,　車両台帳.型式,　車両台帳.車名,　車両台帳.車台番号,"
sql = sql & " 車両台帳.種別,　車両台帳.最大積載量,　車両台帳.車両総重量,　車両台帳.Nox・PM法,　車両台帳.基準緩和事項,　車両台帳.場所, 車両台帳.社用車, 車両台帳.ダンプ"
sql = sql & " from 車両台帳"
sql = sql & " Where (((車両台帳.場所) = " & [Forms]![frm車両台帳]![cmbLoca] & ")"
 sql = sql & " and ((車両台帳.社用車) = " & [Forms]![frm車両台帳]![chkCompCar] & ")"
 sql = sql & " and ((車両台帳.ダンプ) = " & [Forms]![frm車両台帳]![chkDump] & "))"
sql = sql & " order by 車両台帳.番号;"
Me.lstDaicho.RowSource = sql

DoCmd.Requery "lstDaicho"

sql = " select count(車両台帳.登録) AS 登録番号の個数"
sql = sql & " From 車両台帳"
sql = sql & " Where (((車両台帳.場所) = " & [Forms]![frm車両台帳]![cmbLoca] & ")"
 sql = sql & " and ((車両台帳.社用車) = " & [Forms]![frm車両台帳]![chkCompCar] & ")"
 sql = sql & " and ((車両台帳.ダンプ) = " & [Forms]![frm車両台帳]![chkDump] & "))"
 
 
 Dim ds As Database
 Dim rs As Recordset
 
 Set ds = CurrentDb
 Set rs = ds.OpenRecordset(sql)
 
 Me.lblResultRecordCount.Caption = "(" & rs(0) & ")件検索"
 
Me.tbDaichoCarNum = ""
Me.chkCompCar = "0"
End Sub




Private Sub cmbLoca_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDown Then
    cmbLoca.Dropdown
End If

End Sub

Private Sub tbDaichoCarNum_afterupdate()

Dim sql As String
Dim ds As Database
Dim rs As Recordset
Dim rstC As DAO.Recordset

Me.cmbLoca = "17"
Me.chkCompCar = "0"
Me.chkDump = "0"

sql = " select 車両台帳.登録,　車両台帳.番号,　車両台帳.初度登録,　車両台帳.型式,　車両台帳.車名,　車両台帳.車台番号,"
sql = sql & " 車両台帳.種別,　車両台帳.最大積載量,　車両台帳.車両総重量,　車両台帳.Nox・PM法,　車両台帳.基準緩和事項,　車両台帳.場所, 車両台帳.社用車, 車両台帳.ダンプ"
sql = sql & " from 車両台帳"
sql = sql & " Where ((車両台帳.番号) = " & [Forms]![frm車両台帳]![tbDaichoCarNum] & ")"
sql = sql & " order by 車両台帳.番号;"
Me.lstDaicho.RowSource = sql
DoCmd.Requery "lstdaicho"

On Error Resume Next
If Not IsNull(tbDaichoCarNum) Then
Set rstC = CurrentDb.OpenRecordset(sql)
If rstC!社用車.Value = -1 Then
Me.chkCompCar.Value = -1
Else
Me.chkCompCar.Value = 0
End If
If rstC!ダンプ.Value = -1 Then
Me.chkCompCar.Value = -1
Else
Me.chkCompCar.Value = 0
End If

sql = " select count(車両台帳.登録) AS 登録番号の個数"
sql = sql & " From 車両台帳"
sql = sql & " Where ((車両台帳.番号) = " & [Forms]![frm車両台帳]![tbDaichoCarNum] & ")"

 
 Set ds = CurrentDb
 Set rs = ds.OpenRecordset(sql)
 
 Me.lblResultRecordCount.Caption = "(" & rs(0) & ")件検索"
 tbDaichoCarNum.SetFocus
Else

 Me.lblResultRecordCount.Caption = "(0)件検索"
 
 
Exit Sub
tbDaichoCarNum.SetFocus
End If



tbDaichoCarNum.SetFocus
End Sub



