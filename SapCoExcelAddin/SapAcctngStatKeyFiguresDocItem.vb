
Public Class SapAcctngStatKeyFiguresDocItem
    Public item As Dictionary(Of String, SAPCommon.TField)

    Public Sub New()
        item = New Dictionary(Of String, SAPCommon.TField)
    End Sub

    Public Sub SetField(pField As SAPCommon.TField)
        If item.ContainsKey(pField.Name) Then
            item(pField.Name).Value = pField.Value
            item(pField.Name).FType = pField.FType
        Else
            item.Add(pField.Name, pField)
        End If
    End Sub

    Public Sub SetField(pName As String, pValue As String, Optional ByVal pFType As String = "S")
        Dim aField As SAPCommon.TField
        If item.ContainsKey(pName) Then
            item(pName).Value = pValue
            item(pName).FType = pFType
        Else
            aField = New SAPCommon.TField
            aField.SetValues(pName, pValue, pFType)
            item.Add(aField.Name, aField)
        End If
    End Sub
End Class
