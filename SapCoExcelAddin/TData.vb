' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class TData

    Public aTDataDic As Dictionary(Of String, TDataRec)
    Private aIntPar As SAPCommon.TStr
    Private aFieldArray() As String = {}
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Sub New(ByRef pIntPar As SAPCommon.TStr)
        aTDataDic = New Dictionary(Of String, TDataRec)
        aIntPar = pIntPar
    End Sub

    Public Sub addValue(pKey As String, pNAME As String, pVALUE As String, pCURRENCY As String, pFORMAT As String,
                        Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set", Optional pUseAsEmpty As String = "#")
        Dim aTDataRec As TDataRec
        If aTDataDic.ContainsKey(pKey) Then
            aTDataRec = aTDataDic(pKey)
            aTDataRec.setValues(pNAME, pVALUE, pCURRENCY, pFORMAT, pEmty, pEmptyChar, pOperation, pUseAsEmpty)
        Else
            aTDataRec = New TDataRec(aIntPar)
            aTDataRec.setValues(pNAME, pVALUE, pCURRENCY, pFORMAT, pEmty, pEmptyChar, pOperation, pUseAsEmpty)
            aTDataDic.Add(pKey, aTDataRec)
        End If
    End Sub

    Public Sub addValue(pKey As String, pTStrRec As SAPCommon.TStrRec,
                        Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set",
                        Optional pNewStrucname As String = "", Optional pUseAsEmpty As String = "#")
        Dim aTDataRec As TDataRec
        Dim aName As String
        If pNewStrucname <> "" Then
            aName = pNewStrucname & "-" & pTStrRec.Fieldname
        Else
            aName = pTStrRec.Strucname & "-" & pTStrRec.Fieldname
        End If
        If aTDataDic.ContainsKey(pKey) Then
            aTDataRec = aTDataDic(pKey)
            aTDataRec.setValues(aName, pTStrRec.Value, pTStrRec.Currency, pTStrRec.Format, pEmty, pEmptyChar, pOperation, pUseAsEmpty)
        Else
            aTDataRec = New TDataRec(aIntPar)
            aTDataRec.setValues(aName, pTStrRec.Value, pTStrRec.Currency, pTStrRec.Format, pEmty, pEmptyChar, pOperation, pUseAsEmpty)
            aTDataDic.Add(pKey, aTDataRec)
        End If
    End Sub

    Public Sub delData(pKey As String)
        aTDataDic.Remove(pKey)
    End Sub

    Public Function getFirstRecord() As TDataRec
        Dim aTDataRec As TDataRec = Nothing
        Dim aKvb As KeyValuePair(Of String, TDataRec)
        aKvb = aTDataDic.ElementAt(0)
        getFirstRecord = Nothing
        If Not IsNothing(aKvb) Then
            getFirstRecord = aKvb.Value
        End If
    End Function

    Public Sub ws_parse_line_simple(pWsName As String, ByRef pLoff As Integer, i As UInt64, jMax As UInt64, Optional pCoff As Integer = 0, Optional pKey As String = "", Optional pHdrLine As Integer = 1, Optional pUplLine As Integer = 1)
        Dim aDWS As Excel.Worksheet
        Dim aWB As Excel.Workbook
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aDWS = aWB.Worksheets(pWsName)
        Catch Exc As System.Exception
            log.Warn("ws_parse - " & "No " & pWsName & " Sheet in current workbook.")
            MsgBox("No " & pWsName & " Sheet in current workbook. Check the WS Parameters",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap LTP")
            Exit Sub
        End Try
        Dim aName As String = ""
        Dim aUpl As String = ""
        Dim j As Integer
        Dim k As Integer
        Dim aKey As String
        If pKey = "" Or CStr(aDWS.Cells(i, 1).value) = pKey Then
            aKey = CStr(i)
            k = 1
            For j = pCoff + 1 To jMax
                aName = CStr(aDWS.Cells(pHdrLine, j).value)
                aUpl = CStr(aDWS.Cells(pUplLine, j).value)
                If aName <> "N/A" And aName <> "" And aUpl <> "N" And aUpl <> "" Then
                    addValue(aKey, aName, CStr(aDWS.Cells(i, j).value), CStr(aDWS.Cells(pLoff - 2, j).value), CStr(aDWS.Cells(pLoff - 1, j).value), pEmptyChar:="")
                End If
            Next
        End If
    End Sub

    Public Function setFieldArray(pWs As Excel.Worksheet, pCoff As Integer) As ULong
        ' read the header fields
        Dim j As UInt64 = pCoff + 1
        aFieldArray = {}
        Do
            Array.Resize(aFieldArray, aFieldArray.Length + 1)
            aFieldArray(aFieldArray.Length - 1) = CStr(pWs.Cells(1, j).value)
            j += 1
        Loop While Not String.IsNullOrEmpty(pWs.Cells(1, j).value)
        setFieldArray = j
    End Function

    Public Sub ws_output_line(ByRef pWs As Excel.Worksheet, pDataKey As String, i As UInt64, jMax As UInt64, Optional pCoff As Integer = 0, Optional pClear As Boolean = False, Optional pKey As String = "")
        Dim aRange As Excel.Range
        Dim aTDataRec As TDataRec
        Dim aTStrRec As SAPCommon.TStrRec
        If pClear Then
            aRange = pWs.Range(pWs.Cells(i, pCoff + 1), pWs.Cells(i, jMax))
            aRange.Delete()
        End If
        ' output
        Dim j As UInt64 = pCoff + 1
        Dim aFirst As Boolean = True
        If pDataKey = "" Then
            Dim aKvB_Rec As KeyValuePair(Of String, TDataRec)
            For Each aKvB_Rec In aTDataDic
                aTDataRec = aKvB_Rec.Value
                j = pCoff + 1
                Do
                    If aTDataRec.aTDataRecCol.Contains(CStr(pWs.Cells(1, j).value)) Then
                        aTStrRec = aTDataRec.aTDataRecCol(CStr(pWs.Cells(1, j).value))
                        If aFirst Then
                            pWs.Cells(i, j).value = aTStrRec.formated()
                        Else
                            pWs.Cells(i, j).value = pWs.Cells(i, j).value & ";" & aTStrRec.formated()
                        End If
                    End If
                    j += 1
                Loop While j <= jMax
                aFirst = False
            Next
        Else
            If aTDataDic.ContainsKey(pDataKey) Then
                aTDataRec = aTDataDic(pDataKey)
                Do
                    If aTDataRec.aTDataRecCol.Contains(CStr(pWs.Cells(1, j).value)) Then
                        aTStrRec = aTDataRec.aTDataRecCol(CStr(pWs.Cells(1, j).value))
                        pWs.Cells(i, j).value = aTStrRec.formated()
                    End If
                    j += 1
                Loop While j <= jMax
            End If
        End If
    End Sub

    Public Sub ws_output(pWs As Excel.Worksheet, ByRef pLoff As Integer, pCoff As Integer, pPar As SAPCommon.TStr, Optional pClear As Boolean = True, Optional pKey As String = "")
        Dim aRange As Excel.Range
        Dim i As UInt64 = pLoff + 1
        Dim iMax As UInt64 = i - 1
        Do
            iMax += 1
        Loop While Not String.IsNullOrEmpty(pWs.Cells(iMax, 1).value)
        If pClear Then
            If iMax > i Then
                aRange = pWs.Range(pWs.Cells(i, 1), pWs.Cells(iMax, 1))
                aRange.EntireRow.Delete()
            End If
        Else
            i = iMax
        End If
        ' read the header fields
        Dim j As UInt64 = pCoff + 1
        Dim aFieldArray() As String = {}
        Dim aOutArray() As String = {}
        Do
            Array.Resize(aFieldArray, aFieldArray.Length + 1)
            aFieldArray(aFieldArray.Length - 1) = CStr(pWs.Cells(1, j).value)
            j += 1
        Loop While Not String.IsNullOrEmpty(pWs.Cells(1, j).value)
        ' output
        Dim aKvB_Rec As KeyValuePair(Of String, TDataRec)
        Dim aDataRec As New TDataRec(pPar)
        For Each aKvB_Rec In aTDataDic
            aDataRec = aKvB_Rec.Value
            aOutArray = aDataRec.toArray(aFieldArray)
            aRange = pWs.Range(pWs.Cells(i, 1 + pCoff), pWs.Cells(i, aFieldArray.Length + pCoff))
            aRange.Value = aOutArray
            If Not String.IsNullOrEmpty(pKey) Then
                pWs.Cells(i, 1).value = pKey
            End If
            i += 1
        Next
        pLoff = i - 1
    End Sub

End Class
