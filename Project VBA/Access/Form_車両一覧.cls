VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_車両一覧"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbLoca_Change()

Dim SQL As String

SQL = " select 車両台帳.登録,　車両台帳.番号,　車両台帳.初度登録,　車両台帳.型式,　車両台帳.車名,　車両台帳.車台番号,"
SQL = SQL & " 車両台帳.種別,　車両台帳.最大積載量,　車両台帳.車両総重量,　車両台帳.Nox・PM法,　車両台帳.基準緩和事項,　車両台帳.場所, 車両台帳.社用車, 車両台帳.ダンプ"
SQL = SQL & " from 車両台帳"
SQL = SQL & " Where (((車両台帳.場所) = " & [Forms]![車両一覧]![cmbLoca] & ")"
 SQL = SQL & " and ((車両台帳.社用車) = " & [Forms]![車両一覧]![chkCompCar] & ")"
 SQL = SQL & " and ((車両台帳.ダンプ) = " & [Forms]![車両一覧]![chkDump] & "))"
SQL = SQL & " order by 車両台帳.車台番号;"
Me.lstDaicho.RowSource = SQL

DoCmd.Requery "lstDaicho"

SQL = " select count(車両台帳.登録) AS 登録番号の個数"
SQL = SQL & " From 車両台帳"
SQL = SQL & " Where (((車両台帳.場所) = " & [Forms]![車両一覧]![cmbLoca] & ")"
 SQL = SQL & " and ((車両台帳.社用車) = " & [Forms]![車両一覧]![chkCompCar] & ")"
 SQL = SQL & " and ((車両台帳.ダンプ) = " & [Forms]![車両一覧]![chkDump] & "))"
 
 
 Dim ds As Database
 Dim rs As Recordset
 
 Set ds = CurrentDb
 Set rs = ds.OpenRecordset(SQL)
 
 Me.lblResultRecordCount.Caption = "(" & rs(0) & ")件検索"



End Sub
Private Sub chkcompcar_Change()

Dim SQL As String

SQL = " select 車両台帳.登録,　車両台帳.番号,　車両台帳.初度登録,　車両台帳.型式,　車両台帳.車名,　車両台帳.車台番号,"
SQL = SQL & " 車両台帳.種別,　車両台帳.最大積載量,　車両台帳.車両総重量,　車両台帳.Nox・PM法,　車両台帳.基準緩和事項,　車両台帳.場所, 車両台帳.社用車, 車両台帳.ダンプ"
SQL = SQL & " from 車両台帳"
SQL = SQL & " Where (((車両台帳.場所) = " & [Forms]![車両一覧]![cmbLoca] & ")"
 SQL = SQL & " and ((車両台帳.社用車) = " & [Forms]![車両一覧]![chkCompCar] & ")"
 SQL = SQL & " and ((車両台帳.ダンプ) = " & [Forms]![車両一覧]![chkDump] & "))"
SQL = SQL & " order by 車両台帳.車台番号;"
Me.lstDaicho.RowSource = SQL

DoCmd.Requery "lstDaicho"

SQL = " select count(車両台帳.登録) AS 登録番号の個数"
SQL = SQL & " From 車両台帳"
SQL = SQL & " Where (((車両台帳.場所) = " & [Forms]![車両一覧]![cmbLoca] & ")"
 SQL = SQL & " and ((車両台帳.社用車) = " & [Forms]![車両一覧]![chkCompCar] & ")"
 SQL = SQL & " and ((車両台帳.ダンプ) = " & [Forms]![車両一覧]![chkDump] & "))"
 
 
 Dim ds As Database
 Dim rs As Recordset
 
 Set ds = CurrentDb
 Set rs = ds.OpenRecordset(SQL)
 
 Me.lblResultRecordCount.Caption = "(" & rs(0) & ")件検索"



End Sub
Private Sub chkdump_Change()

Dim SQL As String

SQL = " select 車両台帳.登録,　車両台帳.番号,　車両台帳.初度登録,　車両台帳.型式,　車両台帳.車名,　車両台帳.車台番号,"
SQL = SQL & " 車両台帳.種別,　車両台帳.最大積載量,　車両台帳.車両総重量,　車両台帳.Nox・PM法,　車両台帳.基準緩和事項,　車両台帳.場所, 車両台帳.社用車, 車両台帳.ダンプ"
SQL = SQL & " from 車両台帳"
SQL = SQL & " Where (((車両台帳.場所) = " & [Forms]![車両一覧]![cmbLoca] & ")"
 SQL = SQL & " and ((車両台帳.社用車) = " & [Forms]![車両一覧]![chkCompCar] & ")"
 SQL = SQL & " and ((車両台帳.ダンプ) = " & [Forms]![車両一覧]![chkDump] & "))"
SQL = SQL & " order by 車両台帳.車台番号;"
Me.lstDaicho.RowSource = SQL

DoCmd.Requery "lstDaicho"

SQL = " select count(車両台帳.登録) AS 登録番号の個数"
SQL = SQL & " From 車両台帳"
SQL = SQL & " Where (((車両台帳.場所) = " & [Forms]![車両一覧]![cmbLoca] & ")"
 SQL = SQL & " and ((車両台帳.社用車) = " & [Forms]![車両一覧]![chkCompCar] & ")"
 SQL = SQL & " and ((車両台帳.ダンプ) = " & [Forms]![車両一覧]![chkDump] & "))"
 
 
 Dim ds As Database
 Dim rs As Recordset
 
 Set ds = CurrentDb
 Set rs = ds.OpenRecordset(SQL)
 
 Me.lblResultRecordCount.Caption = "(" & rs(0) & ")件検索"



End Sub

