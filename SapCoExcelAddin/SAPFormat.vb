' Copyright 2017-2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SAPFormat

    Private aWBSMask As SAPCommon.TStr

    Public Sub New(ByRef pIntPar As SAPCommon.TStr)
        Dim aPar As SAPCommon.TStrRec
        Dim aIntParDic As Dictionary(Of String, SAPCommon.TStrRec)
        aIntParDic = pIntPar.getData()
        Dim aKvB As KeyValuePair(Of String, SAPCommon.TStrRec)
        aWBSMask = New SAPCommon.TStr
        For Each aKvB In aIntParDic
            aPar = aKvB.Value
            If aPar.Strucname = "WBS_MASK" Then
                aWBSMask.add(aPar.Strucname, aPar.Fieldname, aPar.Value, "", "")
            End If
        Next
    End Sub

    Public Function unpack(val As String, length As Integer) As String
        Dim ZeroStr As String
        ZeroStr = "000000000000000000000000000000"
        If IsNumeric(val) Then
            unpack = Left(ZeroStr, length - Len(val)) & val
        Else
            unpack = val
        End If
    End Function

    Public Function pspid(val As String, length As Integer) As String
        Dim aVal() As String
        Dim aMaskStr As String
        aVal = val.Split(".")
        If Not String.IsNullOrEmpty(aVal(0)) Then
            If aWBSMask.value("WBS_MASK", aVal(0)) <> "" Then
                aMaskStr = aWBSMask.value("WBS_MASK", aVal(0))
                val = Replace(val, ".", "")
                pspid = val & Right(aMaskStr, Len(aMaskStr) - Len(val))
            Else
                aMaskStr = "000000000000000000000000"
                val = Replace(val, ".", "")
                pspid = val & Left(aMaskStr, length - Len(val))
            End If
        End If
    End Function

    Public Function uneditProj(val As String, length As Integer) As String
        Dim aVal() As String
        Dim aMaskStr As String
        aVal = val.Split(".")
        If Not String.IsNullOrEmpty(aVal(0)) Then
            If aWBSMask.value("WBS_MASK", aVal(0)) <> "" Then
                aMaskStr = aWBSMask.value("WBS_MASK", aVal(0))
                val = Replace(val, ".", "")
                uneditProj = val & Right(aMaskStr, Len(aMaskStr) - Len(val))
            Else
                aMaskStr = "000000000000000000000000"
                val = Replace(val, ".", "")
                uneditProj = val & Left(aMaskStr, length - Len(val))
            End If
        End If
    End Function
End Class
