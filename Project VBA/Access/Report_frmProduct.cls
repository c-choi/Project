VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_frmProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public Loca As Variant


Private Sub Report_Load()

Dim recordcount As Integer

lblLoca.Caption = "Bugui Inventory"


If Location = -1 Then
lblLoca.Caption = "Itaewon"
End If
If Location = 0 Then
lblLoca.Caption = "Myeongdong"
End If
If Loca = "Bugui Inventory" Then
lblLoca.Caption = "Bugui Inventory"
End If


End Sub


