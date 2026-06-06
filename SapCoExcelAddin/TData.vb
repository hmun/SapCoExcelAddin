Imports SAP.Middleware.Connector

Public Class TData

    Public aTDataDic As Dictionary(Of String, TDataRec)
    Private aPar As SAPCommon.TStr
    Private aFieldArray() As String = {}
    Private aIsValueArray() As String = {}
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Sub New(ByRef pPar As SAPCommon.TStr)
        aTDataDic = New Dictionary(Of String, TDataRec)
        aPar = pPar
    End Sub

    Public Sub addValue(pKey As String, pNAME As String, pVALUE As String, pCURRENCY As String, pFORMAT As String,
                        Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set")
        Dim aTDataRec As TDataRec
        If aTDataDic.ContainsKey(pKey) Then
            aTDataRec = aTDataDic(pKey)
            aTDataRec.setValues(pNAME, pVALUE, pCURRENCY, pFORMAT, pEmty, pEmptyChar, pOperation)
        Else
            aTDataRec = New TDataRec(aPar)
            aTDataRec.setValues(pNAME, pVALUE, pCURRENCY, pFORMAT, pEmty, pEmptyChar, pOperation)
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
            aTDataRec = New TDataRec(aPar)
            aTDataRec.setValues(aName, pTStrRec.Value, pTStrRec.Currency, pTStrRec.Format, pEmty, pEmptyChar, pOperation, pUseAsEmpty)
            aTDataDic.Add(pKey, aTDataRec)
        End If
    End Sub

    Public Sub addValue(pKey As String, ByRef oStruc As IRfcStructure, Optional pStrucName As String = "")
        If Not oStruc Is Nothing Then
            Dim aStrucName As String = If(pStrucName = "", oStruc.Metadata.Name, pStrucName)
            For j As Integer = 0 To oStruc.Count - 1
                addValue(pKey, aStrucName & "-" & oStruc(j).Metadata.Name, CStr(oStruc(j).GetValue), "", "")
            Next
        End If
    End Sub

    Public Sub addValues(ByRef oTable As IRfcTable, Optional pStrucName As String = "")
        Dim oStruc As IRfcStructure = Nothing
        Dim aStrucName As String
        If Not oTable Is Nothing Then
            aStrucName = If(pStrucName = "", oTable(0).Metadata.Name, pStrucName)
            For i As Integer = 0 To oTable.Count - 1
                addValue(CStr(i), oTable(i), aStrucName)
            Next
        End If
    End Sub

    Public Sub addValue(pKey As String, pTStrRec As SAPCommon.TStrRec,
                        Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set",
                        Optional pNewStrucname As String = "")
        Dim aTDataRec As TDataRec
        Dim aName As String
        If pNewStrucname <> "" Then
            aName = pNewStrucname & "-" & pTStrRec.Fieldname
        Else
            aName = pTStrRec.Strucname & "-" & pTStrRec.Fieldname
        End If
        If aTDataDic.ContainsKey(pKey) Then
            aTDataRec = aTDataDic(pKey)
            aTDataRec.setValues(aName, pTStrRec.Value, pTStrRec.Currency, pTStrRec.Format, pEmty, pEmptyChar, pOperation)
        Else
            aTDataRec = New TDataRec(aPar)
            aTDataRec.setValues(aName, pTStrRec.Value, pTStrRec.Currency, pTStrRec.Format, pEmty, pEmptyChar, pOperation)
            aTDataDic.Add(pKey, aTDataRec)
        End If
    End Sub

    Public Sub delData(pKey As String)
        aTDataDic.Remove(pKey)
    End Sub

    Public Function getPostingRecord() As TDataRec
        Dim aTDataRec As TDataRec = Nothing
        Dim aKvb As KeyValuePair(Of String, TDataRec)
        For Each aKvb In aTDataDic
            aTDataRec = aKvb.Value
            If aTDataRec.getPost(aPar) <> "" Then
                getPostingRecord = aTDataRec
                Exit Function
            End If
        Next
        getPostingRecord = Nothing
    End Function

    Public Function getFirstRecord() As TDataRec
        Dim aTDataRec As TDataRec = Nothing
        Dim aKvb As KeyValuePair(Of String, TDataRec)
        aKvb = aTDataDic.ElementAt(0)
        getFirstRecord = Nothing
        If Not IsNothing(aKvb) Then
            getFirstRecord = aKvb.Value
        End If
    End Function

    Public Sub ws_parse_line_simple(ByRef pWs As Excel.Worksheet, ByRef pLoff As Integer, i As UInt64, jMax As UInt64, Optional pCoff As Integer = 0, Optional pKey As String = "", Optional pHdrLine As Integer = 1)
        Dim aEmptyChar As String = If(aPar.value("CHAR", "EMPTY") <> "", aPar.value("CHAR", "EMPTY"), "#")
        Dim aName As String = ""
        Dim j As Integer
        Dim k As Integer
        Dim aKey As String
        If pKey = "" Or CStr(pWs.Cells(i, 1).value) = pKey Then
            aKey = CStr(i)
            k = 1
            For j = pCoff + 1 To jMax
                aName = CStr(pWs.Cells(pHdrLine, j).value)
                If aName <> "N/A" And aName <> "" Then
                    If CStr(pWs.Cells(i, j).value) = aEmptyChar Then
                        addValue(aKey, aName, "", CStr(pWs.Cells(pLoff - 2, j).value), CStr(pWs.Cells(pLoff - 1, j).value), pEmptyChar:="", pEmty:=True)
                    Else
                        addValue(aKey, aName, CStr(pWs.Cells(i, j).value), CStr(pWs.Cells(pLoff - 2, j).value), CStr(pWs.Cells(pLoff - 1, j).value), pEmptyChar:="", pEmty:=False)
                    End If
                End If
            Next
        End If
    End Sub

    Public Sub ws_parse_line_simple_ps(pWsName As String, ByRef pLoff As Integer, i As UInt64, jMax As UInt64, Optional pCoff As Integer = 0, Optional pKey As String = "", Optional pHdrLine As Integer = 1, Optional pUplLine As Integer = 1)
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

    Public Function getFieldArray(ByRef pWs As Excel.Worksheet, ByRef pFieldArray() As String, ByRef pIsValueArray() As String, pCoff As Integer) As ULong
        ' read the header fields
        Dim j As UInt64 = pCoff + 1
        pFieldArray = {}
        Do
            Array.Resize(pIsValueArray, pIsValueArray.Length + 1)
            pIsValueArray(pIsValueArray.Length - 1) = CStr(pWs.Cells(2, j).value)
            Array.Resize(pFieldArray, pFieldArray.Length + 1)
            pFieldArray(pFieldArray.Length - 1) = CStr(pWs.Cells(1, j).value)
            pFieldArray(pFieldArray.Length - 1) = pFieldArray(pFieldArray.Length - 1).Replace("HEADDATA", "CLIENTDATA")
            pFieldArray(pFieldArray.Length - 1) = pFieldArray(pFieldArray.Length - 1).Replace("EXTENSIONIN", "EXTENSIONOUT")
            j += 1
        Loop While Not String.IsNullOrEmpty(pWs.Cells(1, j).value)
        aFieldArray = pFieldArray
        getFieldArray = j
    End Function

    Public Sub setFieldArray(ByRef pFieldArray() As String, ByRef pIsValueArray() As String)
        aFieldArray = pFieldArray
        aIsValueArray = pIsValueArray
    End Sub

    Public Sub ws_output_line(ByRef pWs As Excel.Worksheet, pDataKey As String, i As UInt64, jMax As UInt64, Optional pCoff As Integer = 0, Optional ByRef pClear As Boolean = False, Optional pKey As String = "")
        Dim aRange As Excel.Range
        ' clear the output
        Dim iMax As UInt64 = i - 1
        Do
            iMax += 1
        Loop While Not String.IsNullOrEmpty(pWs.Cells(iMax, 1).value)
        If pClear Then
            If iMax > i Then
                aRange = pWs.Range(pWs.Cells(i, 1), pWs.Cells(iMax, jMax))
                aRange.EntireRow.Delete()
            End If
            pClear = False
        End If
        ' output
        Dim j As UInt64 = pCoff + 1
        Dim aFirst As Boolean = True
        Dim aDataRec As New TDataRec(aPar)
        If pDataKey = "" Then
            Dim aKvB_Rec As KeyValuePair(Of String, TDataRec)
            For Each aKvB_Rec In aTDataDic
                aDataRec = aKvB_Rec.Value
                aRange = pWs.Range(pWs.Cells(i, 1 + pCoff), pWs.Cells(i, aFieldArray.Length + pCoff))
                aDataRec.toRange(aFieldArray, aIsValueArray, aRange)
                If Not String.IsNullOrEmpty(pKey) Then
                    pWs.Cells(i, 1).value = pKey
                End If
            Next
        Else
            If aTDataDic.ContainsKey(pDataKey) Then
                aDataRec = aTDataDic(pDataKey)
                aRange = pWs.Range(pWs.Cells(i, 1 + pCoff), pWs.Cells(i, aFieldArray.Length + pCoff))
                aDataRec.toRange(aFieldArray, aIsValueArray, aRange)
                If Not String.IsNullOrEmpty(pKey) Then
                    pWs.Cells(i, 1).value = pKey
                End If
            End If
        End If
    End Sub

    Public Function ws_output(ByRef pWs As Excel.Worksheet, pDataKey As String, i As UInt64, jMax As UInt64, Optional pCoff As Integer = 0, Optional ByRef pClear As Boolean = False, Optional pKey As String = "") As UInt64
        Dim aRange As Excel.Range
        Dim aI As UInt64 = 0
        ' clear the output
        Dim iMax As UInt64 = i - 1
        Do
            iMax += 1
        Loop While Not String.IsNullOrEmpty(pWs.Cells(iMax, 1).value)
        If pClear Then
            If iMax > i Then
                aRange = pWs.Range(pWs.Cells(i, 1), pWs.Cells(iMax, jMax))
                aRange.EntireRow.Delete()
            End If
            pClear = False
        End If
        ' output
        Dim aKvB_Rec As KeyValuePair(Of String, TDataRec)
        Dim aDataRec As TDataRec
        For Each aKvB_Rec In aTDataDic
            aDataRec = aKvB_Rec.Value
            aRange = pWs.Range(pWs.Cells(i + aI, 1 + pCoff), pWs.Cells(i + aI, aFieldArray.Length + pCoff))
            aDataRec.toRange(aFieldArray, aIsValueArray, aRange)
            If Not String.IsNullOrEmpty(pKey) Then
                pWs.Cells(i + aI, 1).value = pKey
            End If
            aI += 1
        Next
        ws_output = aI
    End Function

    Public Sub ws_output_ps(pWs As Excel.Worksheet, ByRef pLoff As Integer, pCoff As Integer, pPar As SAPCommon.TStr, Optional pClear As Boolean = True, Optional pKey As String = "")
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
