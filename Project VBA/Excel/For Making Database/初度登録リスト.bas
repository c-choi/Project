Attribute VB_Name = "���x�o�^���X�g"
Sub FirstReg()
''after adding 2 catagories�ő�ύڗ�  �ԗ����d��
Dim FirstReg As Workbook
Dim master As Workbook
Dim Mplatenum As Range
Dim Fplatenum As Range
Dim YearsR As Range
Dim yearsM As Range
Dim RowsC As Integer
Dim Pnum As String

Set master = Workbooks("���C�Y�E�Z�u���}�X�^�t�@�C��.xlsm")
Set FirstReg = Workbooks("20141119 �ۗL�ԗ����x�o�^ ���X�g.xlsx")
master.Activate
Set Mplatenum = master.Sheets(1).Range(Range("i2"), Range("i2").End(xlDown))
RowsC = Mplatenum.Count
Set yearsM = Mplatenum.Offset(0, 23)
FirstReg.Activate
Set Fplatenum = FirstReg.Sheets(1).Range(Range("i5"), Range("i5").End(xlDown))
Set YearsR = Fplatenum.Offset(0, -5)

Application.ScreenUpdating = False
master.Activate

i = 1
j = 1

Do Until i = RowsC + 1

Pnum = Mplatenum.Cells(i).Value
Set findp = Fplatenum.Find(what:=Pnum, lookat:=xlWhole)

If Not findp Is Nothing Then

yearsM.Cells(i).Value = findp.Offset(0, -5)
yearsM.Cells(i).Offset(0, 1).Value = findp.Offset(0, -4)
yearsM.Cells(i).Offset(0, 2).Value = findp.Offset(0, -3)
i = i + 1
Else
i = i + 1
End If
Loop


Application.ScreenUpdating = True
master.Activate
End Sub
