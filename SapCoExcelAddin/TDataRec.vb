' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class TDataRec

    Public aTDataRecCol As Collection
    Private aIntPar As SAPCommon.TStr

    Public Sub New(ByRef pIntPar As SAPCommon.TStr)
        aTDataRecCol = New Collection
        aIntPar = pIntPar
    End Sub

    Public Sub setValues(pNAME As String, pVALUE As String, pCURRENCY As String, pFORMAT As String,
                         Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set", Optional pUseAsEmpty As String = "#")
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNameArray() As String
        Dim aKey As String
        Dim aSTRUCNAME As String = ""
        Dim aFIELDNAME As String = ""
        Dim aValue As String
        If pVALUE = pUseAsEmpty Then
            aValue = " "
        Else
            aValue = pVALUE
            If Not pEmty And aValue = pEmptyChar Then
                Exit Sub
            End If
        End If
        ' do not add empty values

        If InStr(pNAME, "-") <> 0 Then
            aNameArray = Split(pNAME, "-")
            aSTRUCNAME = aNameArray(0)
            aFIELDNAME = aNameArray(1)
        Else
            aSTRUCNAME = ""
            aFIELDNAME = pNAME
        End If
        aKey = pNAME
        If aTDataRecCol.Contains(aKey) Then
            aTStrRec = aTDataRecCol(aKey)
            Select Case pOperation
                Case "add"
                    aTStrRec.addValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
                Case "sub"
                    aTStrRec.subValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
                Case "mul"
                    aTStrRec.mulValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
                Case "div"
                    aTStrRec.divValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
                Case Else
                    aTStrRec.setValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
            End Select
        Else
            aTStrRec = New SAPCommon.TStrRec
            aTStrRec.setValues(aSTRUCNAME, aFIELDNAME, aValue, pCURRENCY, pFORMAT)
            aTDataRecCol.Add(aTStrRec, aKey)
        End If
    End Sub

    Public Sub setValues(pTDataRec As TDataRec, Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set")
        Dim aTStrRec As SAPCommon.TStrRec
        For Each aTStrRec In pTDataRec.aTDataRecCol
            setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmty, pEmptyChar, pOperation)
        Next
    End Sub

    Public Sub addValues(pTDataRec As TDataRec, Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#")
        Dim aTStrRec As SAPCommon.TStrRec
        For Each aTStrRec In pTDataRec.aTDataRecCol
            If aTStrRec.Currency <> "" Then
                setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmty, pEmptyChar, pOperation:="add")
            Else
                setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmty, pEmptyChar, pOperation:="set")
            End If
        Next
    End Sub

    Public Function getColumn(pClmn As String) As SAPCommon.TStrRec
        Dim aTStrRec As SAPCommon.TStrRec
        If aTDataRecCol.Contains(pClmn) Then
            aTStrRec = aTDataRecCol(pClmn)
            getColumn = aTStrRec
        End If
    End Function

    Public Function getWbsRec() As SAPCommon.TStrRec
        Dim aWbsClmn As String = If(aIntPar.value("WBS_COL", "WBS") <> "", aIntPar.value("WBS_COL", "WBS"), "IT_WBS_ELEMENT-WBS_ELEMENT")
        If aTDataRecCol.Contains(aWbsClmn) Then
            getWbsRec = aTDataRecCol(aWbsClmn)
        End If
    End Function

    Public Function getProject() As String
        Dim aWbsClmn As String = If(aIntPar.value("WBS_COL", "PROJECT") <> "", aIntPar.value("WBS_COL", "PROJECT"), "I_PROJECT_DEFINITION")
        Dim aTStrRec As SAPCommon.TStrRec
        getProject = ""
        If aTDataRecCol.Contains(aWbsClmn) Then
            aTStrRec = aTDataRecCol(aWbsClmn)
            getProject = aTStrRec.Value
        End If
    End Function

    Public Function getWbs() As String
        Dim aWbsClmn As String = If(aIntPar.value("WBS_COL", "WBS") <> "", aIntPar.value("WBS_COL", "WBS"), "IT_WBS_ELEMENT-WBS_ELEMENT")
        Dim aTStrRec As SAPCommon.TStrRec
        getWbs = ""
        If aTDataRecCol.Contains(aWbsClmn) Then
            aTStrRec = aTDataRecCol(aWbsClmn)
            getWbs = aTStrRec.Value
        End If
    End Function

    Public Function getProjZZ_REL(pNumber As String) As String
        Dim aClmn As String = If(aIntPar.value("PROJ_COL", "ZZ_REL_" & pNumber) <> "", aIntPar.value("PROJ_COL", "ZZ_REL_" & pNumber), "BAPI_TE_PROJECT_DEFINITION-ZZ_REL_" & pNumber)
        Dim aTStrRec As SAPCommon.TStrRec
        getProjZZ_REL = ""
        If aTDataRecCol.Contains(aClmn) Then
            aTStrRec = aTDataRecCol(aClmn)
            getProjZZ_REL = aTStrRec.Value
        End If
    End Function

    Public Function getWbsZZ_REL(pNumber As String) As String
        Dim aClmn As String = If(aIntPar.value("WBS_COL", "ZZ_REL_" & pNumber) <> "", aIntPar.value("WBS_COL", "ZZ_REL_" & pNumber), "BAPI_TE_WBS_ELEMENT-ZZ_REL_" & pNumber)
        Dim aTStrRec As SAPCommon.TStrRec
        getWbsZZ_REL = ""
        If aTDataRecCol.Contains(aClmn) Then
            aTStrRec = aTDataRecCol(aClmn)
            getWbsZZ_REL = aTStrRec.Value
        End If
    End Function

    Public Function toArray(pFields()) As String()
        Dim aValueArray() As String = {}
        Dim aTStrRec As SAPCommon.TStrRec
        For i = 0 To pFields.Count - 1
            If aTDataRecCol.Contains(pFields(i)) Then
                aTStrRec = aTDataRecCol(pFields(i))
                Array.Resize(aValueArray, aValueArray.Length + 1)
                aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
            End If
        Next
        toArray = aValueArray
    End Function

End Class
