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

Dim SQL As String

SQL = " select 車両台帳.登録,　車両台帳.番号,　車両台帳.初度登録,　車両台帳.型式,　車両台帳.車名,　車両台帳.車台番号,"
SQL = SQL & " 車両台帳.種別,　車両台帳.最大積載量,　車両台帳.車両総重量,　車両台帳.Nox・PM法,　車両台帳.基準緩和事項,　車両台帳.場所, 車両台帳.社用車, 車両台帳.ダンプ"
SQL = SQL & " from 車両台帳"
SQL = SQL & " Where (((車両台帳.場所) = " & [Forms]![frm車両台帳]![cmbLoca] & ")"
 SQL = SQL & " and ((車両台帳.社用車) = " & [Forms]![frm車両台帳]![chkCompCar] & ")"
 SQL = SQL & " and ((車両台帳.ダンプ) = " & [Forms]![frm車両台帳]![chkDump] & "))"
SQL = SQL & " order by 車両台帳.車台番号;"
Me.lstDaicho.RowSource = SQL

DoCmd.Requery "lstDaicho"

SQL = " select count(車両台帳.登録) AS 登録番号の個数"
SQL = SQL & " From 車両台帳"
SQL = SQL & " Where (((車両台帳.場所) = " & [Forms]![frm車両台帳]![cmbLoca] & ")"
 SQL = SQL & " and ((車両台帳.社用車) = " & [Forms]![frm車両台帳]![chkCompCar] & ")"
 SQL = SQL & " and ((車両台帳.ダンプ) = " & [Forms]![frm車両台帳]![chkDump] & "))"
 
 
 Dim ds As Database
 Dim rs As Recordset
 
 Set ds = CurrentDb
 Set rs = ds.OpenRecordset(SQL)
 
 Me.lblResultRecordCount.Caption = "(" & rs(0) & ")件検索"

Me.tbDaichoCarNum = ""

End Sub


Private Sub chkCompCar_afterupdate()

Dim SQL As String

SQL = " select 車両台帳.登録,　車両台帳.番号,　車両台帳.初度登録,　車両台帳.型式,　車両台帳.車名,　車両台帳.車台番号,"
SQL = SQL & " 車両台帳.種別,　車両台帳.最大積載量,　車両台帳.車両総重量,　車両台帳.Nox・PM法,　車両台帳.基準緩和事項,　車両台帳.場所, 車両台帳.社用車, 車両台帳.ダンプ"
SQL = SQL & " from 車両台帳"
SQL = SQL & " Where (((車両台帳.場所) = " & [Forms]![frm車両台帳]![cmbLoca] & ")"
 SQL = SQL & " and ((車両台帳.社用車) = " & [Forms]![frm車両台帳]![chkCompCar] & ")"
 SQL = SQL & " and ((車両台帳.ダンプ) = " & [Forms]![frm車両台帳]![chkDump] & "))"
SQL = SQL & " order by 車両台帳.車台番号;"
Me.lstDaicho.RowSource = SQL

DoCmd.Requery "lstDaicho"

SQL = " select count(車両台帳.登録) AS 登録番号の個数"
SQL = SQL & " From 車両台帳"
SQL = SQL & " Where (((車両台帳.場所) = " & [Forms]![frm車両台帳]![cmbLoca] & ")"
 SQL = SQL & " and ((車両台帳.社用車) = " & [Forms]![frm車両台帳]![chkCompCar] & ")"
 SQL = SQL & " and ((車両台帳.ダンプ) = " & [Forms]![frm車両台帳]![chkDump] & "))"
 
 
 Dim ds As Database
 Dim rs As Recordset
 
 Set ds = CurrentDb
 Set rs = ds.OpenRecordset(SQL)
 
 Me.lblResultRecordCount.Caption = "(" & rs(0) & ")件検索"

Me.tbDaichoCarNum = ""
Me.chkDump = "0"
End Sub


Private Sub chkDump_afterupdate()

Dim SQL As String

SQL = " select 車両台帳.登録,　車両台帳.番号,　車両台帳.初度登録,　車両台帳.型式,　車両台帳.車名,　車両台帳.車台番号,"
SQL = SQL & " 車両台帳.種別,　車両台帳.最大積載量,　車両台帳.車両総重量,　車両台帳.Nox・PM法,　車両台帳.基準緩和事項,　車両台帳.場所, 車両台帳.社用車, 車両台帳.ダンプ"
SQL = SQL & " from 車両台帳"
SQL = SQL & " Where (((車両台帳.場所) = " & [Forms]![frm車両台帳]![cmbLoca] & ")"
 SQL = SQL & " and ((車両台帳.社用車) = " & [Forms]![frm車両台帳]![chkCompCar] & ")"
 SQL = SQL & " and ((車両台帳.ダンプ) = " & [Forms]![frm車両台帳]![chkDump] & "))"
SQL = SQL & " order by 車両台帳.車台番号;"
Me.lstDaicho.RowSource = SQL

DoCmd.Requery "lstDaicho"

SQL = " select count(車両台帳.登録) AS 登録番号の個数"
SQL = SQL & " From 車両台帳"
SQL = SQL & " Where (((車両台帳.場所) = " & [Forms]![frm車両台帳]![cmbLoca] & ")"
 SQL = SQL & " and ((車両台帳.社用車) = " & [Forms]![frm車両台帳]![chkCompCar] & ")"
 SQL = SQL & " and ((車両台帳.ダンプ) = " & [Forms]![frm車両台帳]![chkDump] & "))"
 
 
 Dim ds As Database
 Dim rs As Recordset
 
 Set ds = CurrentDb
 Set rs = ds.OpenRecordset(SQL)
 
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

Dim SQL As String
Dim ds As Database
Dim rs As Recordset
Dim rstC As DAO.Recordset



SQL = " select 車両台帳.登録,　車両台帳.番号,　車両台帳.初度登録,　車両台帳.型式,　車両台帳.車名,　車両台帳.車台番号,"
SQL = SQL & " 車両台帳.種別,　車両台帳.最大積載量,　車両台帳.車両総重量,　車両台帳.Nox・PM法,　車両台帳.基準緩和事項,　車両台帳.場所, 車両台帳.社用車, 車両台帳.ダンプ"
SQL = SQL & " from 車両台帳"
SQL = SQL & " Where ((車両台帳.番号) = " & [Forms]![frm車両台帳]![tbDaichoCarNum] & ")"
Me.lstDaicho.RowSource = SQL
DoCmd.Requery "lstdaicho"

On Error Resume Next
If Not IsNull(tbDaichoCarNum) Then
Set rstC = CurrentDb.OpenRecordset(SQL)
If rstC!社用車.Value = -1 Then
Me.chkCompCar.Value = -1
Else
Me.chkCompCar.Value = 0
End If

SQL = " select count(車両台帳.登録) AS 登録番号の個数"
SQL = SQL & " From 車両台帳"
SQL = SQL & " Where ((車両台帳.番号) = " & [Forms]![frm車両台帳]![tbDaichoCarNum] & ")"

 
 Set ds = CurrentDb
 Set rs = ds.OpenRecordset(SQL)
 
 Me.lblResultRecordCount.Caption = "(" & rs(0) & ")件検索"
 tbDaichoCarNum.SetFocus
Else

 Me.lblResultRecordCount.Caption = "(0)件検索"
 
 
Exit Sub
tbDaichoCarNum.SetFocus
End If

Me.cmbLoca = "17"
Me.chkDump = "0"

tbDaichoCarNum.SetFocus
End Sub



