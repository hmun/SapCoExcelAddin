Public Class TPostingData

    Public aTPostingDataDic As Dictionary(Of String, TPostingDataRec)
    Private aPar As SAPCommon.TStr

    Public Sub New(ByRef pPar As SAPCommon.TStr)
        aTPostingDataDic = New Dictionary(Of String, TPostingDataRec)
        aPar = pPar
    End Sub

    Public Sub addTPostingDataRec(pKey As String, pTPostingDataRec As TPostingDataRec, Optional pNr As String = "")
        Dim aTPostingDataRec As TPostingDataRec
        If aTPostingDataDic.ContainsKey(pKey) Then
            aTPostingDataRec = aTPostingDataDic(pKey)
            aTPostingDataRec.addAmounts(pTPostingDataRec)
        Else
            aTPostingDataRec = New TPostingDataRec(aPar, pNr:=pNr)
            aTPostingDataRec.addValues(pTPostingDataRec)
            aTPostingDataDic.Add(pKey, aTPostingDataRec)
        End If
    End Sub

    Public Function newTPostingDataRec(ByRef pDic As Dictionary(Of String, SAPCommon.TField), Optional pEmpty As Boolean = False, Optional pEmptyChar As String = "#", Optional pNr As String = "") As TPostingDataRec
        newTPostingDataRec = New TPostingDataRec(pDic:=pDic, pPar:=aPar, pEmpty:=pEmpty, pEmptyChar:=pEmptyChar, pNr:=pNr)
    End Function

    Public Sub addValue(pKey As String, pNAME As String, pVALUE As String, pCURRENCY As String, pFORMAT As String,
                        Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set", Optional pNr As String = "")
        Dim aTPostingDataRec As TPostingDataRec
        If aTPostingDataDic.ContainsKey(pKey) Then
            aTPostingDataRec = aTPostingDataDic(pKey)
            aTPostingDataRec.setValues(pNAME, pVALUE, pCURRENCY, pFORMAT, pEmty, pEmptyChar, pOperation)
        Else
            aTPostingDataRec = New TPostingDataRec(aPar, pNr:=pNr)
            aTPostingDataRec.setValues(pNAME, pVALUE, pCURRENCY, pFORMAT, pEmty, pEmptyChar, pOperation)
            aTPostingDataDic.Add(pKey, aTPostingDataRec)
        End If
    End Sub

    Public Sub addValue(pKey As String, pTStrRec As SAPCommon.TStrRec,
                        Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set",
                        Optional pNewStrucname As String = "", Optional pNr As String = "")
        Dim aTPostingDataRec As TPostingDataRec
        Dim aName As String
        If pNewStrucname <> "" Then
            aName = pNewStrucname & "-" & pTStrRec.Fieldname
        Else
            aName = pTStrRec.Strucname & "-" & pTStrRec.Fieldname
        End If
        If aTPostingDataDic.ContainsKey(pKey) Then
            aTPostingDataRec = aTPostingDataDic(pKey)
            aTPostingDataRec.setValues(aName, pTStrRec.Value, pTStrRec.Currency, pTStrRec.Format, pEmty, pEmptyChar, pOperation)
        Else
            aTPostingDataRec = New TPostingDataRec(aPar, pNr:=pNr)
            aTPostingDataRec.setValues(aName, pTStrRec.Value, pTStrRec.Currency, pTStrRec.Format, pEmty, pEmptyChar, pOperation)
            aTPostingDataDic.Add(pKey, aTPostingDataRec)
        End If
    End Sub

    Public Sub delData(pKey As String)
        aTPostingDataDic.Remove(pKey)
    End Sub

End Class
