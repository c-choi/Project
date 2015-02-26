VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm車両一覧"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmblstComp_AfterUpdate()

Dim SQL As String

SQL = " select 車両台帳.登録, 車両台帳.番号, 車両台帳.車台番号, 車両台帳.所有者, 車両台帳.[リ-ス満了日], 車両台帳.[月額リ-ス料], "
SQL = SQL & " 車両台帳.初度登録,　車両台帳.平成24年, 車両台帳.平成25年, 車両台帳.平成26年, 車両台帳.平成27年, 車両台帳.形状 , 車両台帳.型式,"
SQL = SQL & " 車両台帳.最大積載量, 車両台帳.Nox・PM法,　車両台帳.基準緩和事項,　車両台帳.使用者"
SQL = SQL & " from 車両台帳"
SQL = SQL & " Where ((車両台帳.使用者) ='" & [Forms]![frm車両一覧]![cmblstComp] & "')"
SQL = SQL & " order by 車両台帳.車台番号"
Me.lstFirstReg.RowSource = SQL

DoCmd.Requery "lstFirstReg"

SQL = " select count(車両台帳.車台番号) AS 車台番号の個数"
SQL = SQL & " From 車両台帳"
SQL = SQL & " Where ((車両台帳.使用者) = " & "'" & [Forms]![frm車両一覧]![cmblstComp] & "'" & ") "

 
 Dim ds As Database
 Dim rs As Recordset
 
 Set ds = CurrentDb
 Set rs = ds.OpenRecordset(SQL)
 
 Me.lblResultRecordcount2.Caption = "(" & rs(0) & ")件検索"



End Sub

Private Sub cmblstComp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
     cmblstComp.Dropdown
    End If
    
End Sub


