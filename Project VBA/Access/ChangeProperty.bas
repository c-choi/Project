Attribute VB_Name = "ChangeProperty"
Function ChangeProperty(strPropname As String, varProptype As Variant, varPropvalue As Variant) As Integer

Dim dbs As Object, prp As Variant
Const conPropNotFoundError = 3270

Set dbs = CurrentDb
On Error GoTo change_err
dbs.Properties(strPropname) = varPropvalue
ChangeProperty = True

change_bye:
Exit Function

change_err:
    If Err = conPropNotFoundError Then
        Set prp = dbs.CreateProperty(strPropname, _
            varProptype, varPropvalue)
        dbs.Properties.Append prp
        Resume Next
    Else
    
        ChangeProperty = False
        Resume change_bye
    End If

End Function
