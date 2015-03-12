VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public Loca As Variant
Public comID As Variant



Private Sub cmdAddColor_Click()

Dim KindID As String
On Error GoTo cmdAddcolor_Click_Err
    
    DoCmd.OpenForm "frmaddcolor", acNormal, , , acFormAdd, acDialog


    

    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

cmdAddcolor_Click_Exit:
    Exit Sub

cmdAddcolor_Click_Err:
    MsgBox Error$
    Resume cmdAddcolor_Click_Exit


End Sub

Private Sub cmdClose_Click()
comID = cmdkind
DoCmd.Close acForm, Me.Name

End Sub

Private Sub form_DataChange(ByVal Reason As Long)


Dim sql As String

If Me.CmbLeather.Value = "Crocodile" Then

sql = " SELECT colorList.ID, IIf([MatShiny]=0,Replace([MatShiny],0,'Mat'),Replace([MatShiny],-1,'Shiny')) AS [Mat/Shiny],"
sql = sql & " colorList.Color, IIf([Nubk]=0,Replace([Nubk],0,''),Replace([Nubk],-1,'Nubuck')) AS Nubuck"
sql = sql & " FROM colorList ORDER BY IIf([MatShiny]=0,Replace([MatShiny],0,'Mat'),Replace([MatShiny],-1,'Shiny')) DESC , colorList.Color, colorlist.nubk;"

Me.lstColor.RowSource = sql
Me.lstColor.Requery

Else

sql = " SELECT colorList.ID, IIf([MatShiny]=0,Replace([MatShiny],0,'Mat'),Replace([MatShiny],-1,'Shiny')) AS [Mat/Shiny],"
sql = sql & " colorList.Color, IIf([Nubk]=0,Replace([Nubk],0,''),Replace([Nubk],-1,'Nubuck')) AS Nubuck"
sql = sql & " FROM colorList WHERE (((colorList.crocoonly)=0))"
sql = sql & " ORDER BY IIf([MatShiny]=0,Replace([MatShiny],0,'Mat'),Replace([MatShiny],-1,'Shiny')) DESC , colorList.Color;"

Me.lstColor.RowSource = sql
Me.lstColor.Requery
End If
End Sub


Private Sub CmbLeather_AfterUpdate()
Dim sql As String

If Me.CmbLeather.Value = "Crocodile" Then

sql = " SELECT colorList.ID, IIf([MatShiny]=0,Replace([MatShiny],0,'Mat'),Replace([MatShiny],-1,'Shiny')) AS [Mat/Shiny],"
sql = sql & " colorList.Color, IIf([Nubk]=0,Replace([Nubk],0,''),Replace([Nubk],-1,'Nubuck')) AS Nubuck"
sql = sql & " FROM colorList ORDER BY IIf([MatShiny]=0,Replace([MatShiny],0,'Mat'),Replace([MatShiny],-1,'Shiny')) DESC , colorList.Color, colorlist.nubk;"

Me.lstColor.RowSource = sql
Me.lstColor.Requery

Else

sql = " SELECT colorList.ID, IIf([MatShiny]=0,Replace([MatShiny],0,'Mat'),Replace([MatShiny],-1,'Shiny')) AS [Mat/Shiny],"
sql = sql & " colorList.Color, IIf([Nubk]=0,Replace([Nubk],0,''),Replace([Nubk],-1,'Nubuck')) AS Nubuck"
sql = sql & " FROM colorList WHERE (((colorList.crocoonly)=0))"
sql = sql & " ORDER BY IIf([MatShiny]=0,Replace([MatShiny],0,'Mat'),Replace([MatShiny],-1,'Shiny')) DESC , colorList.Color;"

Me.lstColor.RowSource = sql
Me.lstColor.Requery
End If
End Sub
'------------------------------------------------------------
' cmdNew_Click
'
'------------------------------------------------------------
Private Sub cmdNew_Click()
On Error GoTo cmdNew_Click_Err

    On Error Resume Next
    DoCmd.GoToRecord , "", acNewRec
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If


cmdNew_Click_Exit:
    Exit Sub

cmdNew_Click_Err:
    MsgBox Error$
    Resume cmdNew_Click_Exit

End Sub


'------------------------------------------------------------
' cmdDelete_Click
'
'------------------------------------------------------------
Private Sub cmdDelete_Click()
On Error GoTo cmdDelete_Click_Err

    On Error Resume Next
    DoCmd.GoToControl Screen.PreviousControl.Name
    Err.Clear
    If (Not Form.NewRecord) Then
        DoCmd.RunCommand acCmdDeleteRecord
    End If
    If (Form.NewRecord And Not Form.Dirty) Then
        Beep
    End If
    If (Form.NewRecord And Form.Dirty) Then
        DoCmd.RunCommand acCmdUndo
    End If
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If


cmdDelete_Click_Exit:
    Exit Sub

cmdDelete_Click_Err:
    MsgBox Error$
    Resume cmdDelete_Click_Exit

End Sub


'------------------------------------------------------------
' cmdGotoFirst_Click
'
'------------------------------------------------------------
Private Sub cmdGotoFirst_Click()
On Error GoTo cmdGotoFirst_Click_Err

    DoCmd.GoToRecord , "", acFirst


cmdGotoFirst_Click_Exit:
    Exit Sub

cmdGotoFirst_Click_Err:
    MsgBox Error$
    Resume cmdGotoFirst_Click_Exit

End Sub


'------------------------------------------------------------
' cmdgotoLast_Click
'
'------------------------------------------------------------
Private Sub cmdgotoLast_Click()
On Error GoTo cmdgotoLast_Click_Err

    DoCmd.GoToRecord , "", acLast


cmdgotoLast_Click_Exit:
    Exit Sub

cmdgotoLast_Click_Err:
    MsgBox Error$
    Resume cmdgotoLast_Click_Exit

End Sub


'------------------------------------------------------------
' cmdgotoprev_Click
'
'------------------------------------------------------------
Private Sub cmdgotoprev_Click()
On Error GoTo cmdgotoprev_Click_Err

    On Error Resume Next
    DoCmd.GoToRecord , "", acPrevious
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If


cmdgotoprev_Click_Exit:
    Exit Sub

cmdgotoprev_Click_Err:
    MsgBox Error$
    Resume cmdgotoprev_Click_Exit

End Sub


'------------------------------------------------------------
' cmdgotoNext_Click
'
'------------------------------------------------------------
Private Sub cmdgotoNext_Click()
On Error GoTo cmdgotoNext_Click_Err

    On Error Resume Next
    DoCmd.GoToRecord , "", acNext
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If


cmdgotoNext_Click_Exit:
    Exit Sub

cmdgotoNext_Click_Err:
    MsgBox Error$
    Resume cmdgotoNext_Click_Exit

End Sub


'------------------------------------------------------------
' cmdPreview_Click
'
'------------------------------------------------------------
Private Sub cmdPreview_Click()
On Error GoTo cmdPreview_Click_Err

    If (TogLoca.Caption = "Myeongdong") Then
            Loca = "Myeongdong"
        DoCmd.OpenReport "frmProduct", acViewPreview, "", "[ProductList]![Location]=0", acDialog

    End If
    If (TogLoca.Caption = "Itaewon") Then
            Loca = "Itaewon"
        DoCmd.OpenReport "frmProduct", acViewPreview, "", "[ProductList]![Location]=-1", acDialog

    End If
    If (TogLoca.Caption = "All") Then
            Loca = "Bugui Inventory"

        DoCmd.OpenReport "frmProduct", acViewPreview, "", , acDialog
            
    End If
    
cmdPreview_Click_Exit:
    Exit Sub

cmdPreview_Click_Err:
    MsgBox Error$
    Resume cmdPreview_Click_Exit

End Sub

Private Sub Togloca_Click()
If TogLoca.Caption = "All" Then
TogLoca.Caption = "Myeongdong"
Loca = "Myeongdong"
ElseIf TogLoca.Caption = "Myeongdong" Then
TogLoca.Caption = "Itaewon"
Loca = "Itaewon"

ElseIf TogLoca.Caption = "Itaewon" Then
TogLoca.Caption = "All"
Loca = "Bugui Inventory"
End If

End Sub

Private Sub cmbLeather_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDown Then
    CmbLeather.Dropdown
End If

End Sub


Private Sub cmbGrade_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDown Then
    cmbGrade.Dropdown
End If

End Sub


Private Sub cmbQuan_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDown Then
    cmbQuan.Dropdown
End If

End Sub

Private Sub cmbkind_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDown Then
    cmbkind.Dropdown
End If

End Sub

'------------------------------------------------------------
' cmdAddKind_Click
'
'------------------------------------------------------------
Private Sub cmdAddKind_Click()
Dim KindID As String
If cmbkind.Value <> "" Then
KindID = cmbkind.Value
Else
cmbkind = ""
End If
On Error GoTo cmdAddKind_Click_Err
    
    DoCmd.OpenForm "frmaddproductkind", acNormal, , , acFormAdd, acDialog


    

    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

cmdAddKind_Click_Exit:
    Exit Sub

cmdAddKind_Click_Err:
    MsgBox Error$
    Resume cmdAddKind_Click_Exit

End Sub

Private Sub cmdLeather_Click()
Dim LeatherID As String
If CmbLeather.Value <> "" Then
LeatherID = CmbLeather.Value
Else
CmbLeather = ""
End If
On Error GoTo cmdleather_Click_Err
    
    DoCmd.OpenForm "frmaddLeather", acNormal, , , acFormAdd, acDialog


    

    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

cmdLeather_Click_Exit:
    Exit Sub

cmdleather_Click_Err:
    MsgBox Error$
    Resume cmdLeather_Click_Exit

End Sub

