Public Class TPostingDataRec

    Public aTPostingDataRecCol As Collection
    Private aNonKeyArray() As String = {}
    Private aValueArray() As String = {}
    Private aNr As String = ""

    Public Sub New(ByRef pPar As SAPCommon.TStr, Optional pNr As String = "")
        aTPostingDataRecCol = New Collection
        aNr = pNr
        initArrays(pPar)
    End Sub

    Public Sub New(ByRef pDic As Dictionary(Of String, SAPCommon.TField), ByRef pPar As SAPCommon.TStr, Optional pEmpty As Boolean = False, Optional pEmptyChar As String = "#", Optional pNr As String = "")
        Dim aKvb As KeyValuePair(Of String, SAPCommon.TField)
        Dim aTField As SAPCommon.TField
        aNr = pNr
        initArrays(pPar)
        aTPostingDataRecCol = New Collection
        For Each aKvb In pDic
            aTField = aKvb.Value
            setValues(aTField.Name, aTField.Value, "", aTField.FType, pEmpty:=pEmpty, pEmptyChar:=pEmptyChar)
        Next
    End Sub

    Private Function initArrays(ByRef pPar As SAPCommon.TStr)
        Dim aNonKeyFields As String = If(pPar.value("GEN" & aNr, "NON_KEY_STR") <> "", CStr(pPar.value("GEN" & aNr, "NON_KEY_STR")), "")
        If Not String.IsNullOrEmpty(aNonKeyFields) Then
            aNonKeyArray = aNonKeyFields.Split(",")
        End If
        Dim aValueFields As String = If(pPar.value("GEN" & aNr, "VALUE_STR") <> "", CStr(pPar.value("GEN" & aNr, "VALUE_STR")), "Amount")
        If Not String.IsNullOrEmpty(aNonKeyFields) Then
            aValueArray = aValueFields.Split(",")
        End If

    End Function

    Public Sub setValues(pNAME As String, pVALUE As String, pCURRENCY As String, pFORMAT As String,
                         Optional pEmpty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set")
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNameArray() As String
        Dim aKey As String
        Dim aSTRUCNAME As String = ""
        Dim aFIELDNAME As String = ""
        ' do not add empty values
        If Not pEmpty And pVALUE = pEmptyChar Then
            Exit Sub
        End If
        If InStr(pNAME, "-") <> 0 Then
            aNameArray = Split(pNAME, "-")
            aSTRUCNAME = aNameArray(0)
            aFIELDNAME = aNameArray(1)
        Else
            aSTRUCNAME = ""
            aFIELDNAME = pNAME
        End If
        aKey = pNAME
        If aTPostingDataRecCol.Contains(aKey) Then
            aTStrRec = aTPostingDataRecCol(aKey)
            Select Case pOperation
                Case "add"
                    aTStrRec.addValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
                Case "sub"
                    aTStrRec.subValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
                Case "mul"
                    aTStrRec.mulValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
                Case "div"
                    aTStrRec.divValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
                Case Else
                    aTStrRec.setValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
            End Select
        Else
            aTStrRec = New SAPCommon.TStrRec
            aTStrRec.setValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
            aTPostingDataRecCol.Add(aTStrRec, aKey)
        End If
    End Sub

    Public Sub setValues(pTPostingDataRec As TPostingDataRec, Optional pEmpty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set")
        Dim aTStrRec As SAPCommon.TStrRec
        For Each aTStrRec In pTPostingDataRec.aTPostingDataRecCol
            setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmpty, pEmptyChar, pOperation)
        Next
    End Sub

    Public Sub addAmounts(pTPostingDataRec As TPostingDataRec, Optional pEmpty As Boolean = False, Optional pEmptyChar As String = "#")
        Dim aTStrRec As SAPCommon.TStrRec
        For Each aTStrRec In pTPostingDataRec.aTPostingDataRecCol
            If isValue(aTStrRec.getKey()) Then
                setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmpty, pEmptyChar, pOperation:="add")
            End If
        Next
    End Sub

    Public Sub addValues(pTPostingDataRec As TPostingDataRec, Optional pEmpty As Boolean = False, Optional pEmptyChar As String = "#")
        Dim aTStrRec As SAPCommon.TStrRec
        For Each aTStrRec In pTPostingDataRec.aTPostingDataRecCol
            If isValue(aTStrRec.getKey()) Then
                setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmpty, pEmptyChar, pOperation:="add")
            Else
                setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmpty, pEmptyChar, pOperation:="set")
            End If
        Next
    End Sub

    Public Function getKey() As String
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aKey As String = ""
        For Each aTStrRec In aTPostingDataRecCol
            If isKey(aTStrRec.getKey()) Then
                If Not String.IsNullOrEmpty(aKey) Then
                    aKey = aKey & "_"
                End If
                aKey = aKey & aTStrRec.getKey() & "|" & aTStrRec.Value
            End If
        Next
        getKey = aKey
    End Function

    Public Function isZero() As Boolean
        Dim aTStrRec As SAPCommon.TStrRec
        isZero = True
        For Each aTStrRec In aTPostingDataRecCol
            If isValue(aTStrRec.getKey()) Then
                If CDbl(aTStrRec.Value) <> 0 Then
                    isZero = False
                    Exit For
                End If
            End If
        Next
    End Function

    Public Function isValue(pName As String) As Boolean
        isValue = False
        Dim count As Integer
        For count = 0 To aValueArray.Length - 1
            If pName.Contains(aValueArray(count)) Then
                isValue = True
                Exit For
            End If
        Next
    End Function

    Public Function isKey(pName As String) As Boolean
        isKey = True
        Dim count As Integer
        For count = 0 To aNonKeyArray.Length - 1
            If pName.Contains(aNonKeyArray(count)) Then
                isKey = False
                Exit For
            End If
        Next
    End Function

End Class
