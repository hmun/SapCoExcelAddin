' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class TSAP_WbsGenData

    Public aHdrRec As TDataRec
    Public aDataDic As TDataDic

    Public aStrucDic As Dictionary(Of String, RfcStructureMetadata)
    Public aParamDic As Dictionary(Of String, RfcParameterMetadata)

    Private aPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private aFunction As String
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Sub New(ByRef pPar As SAPCommon.TStr, ByRef pIntPar As SAPCommon.TStr, ByRef pSAPWBSPI As SAPWBSPI, pFunction As String)
        aPar = pPar
        aIntPar = pIntPar
        aFunction = pFunction
        aDataDic = New TDataDic(aIntPar)
        aHdrRec = New TDataRec(aIntPar)
        aStrucDic = New Dictionary(Of String, RfcStructureMetadata)
        aParamDic = New Dictionary(Of String, RfcParameterMetadata)
        If pFunction = "GetStatus" Then
            pSAPWBSPI.getMeta_GetStatus(aParamDic, aStrucDic)
        ElseIf pFunction = "SetStatus" Then
            pSAPWBSPI.getMeta_SetStatus(aParamDic, aStrucDic)
        ElseIf pFunction = "GetData" Then
            pSAPWBSPI.getMeta_GetData(aParamDic, aStrucDic)
        End If
    End Sub

    ' not needed as the WBS Status BAPIs do not have header fields but left here for completenes
    Public Function fillHeader(pData As TData) As Boolean
        Dim aKvb As KeyValuePair(Of String, SAPCommon.TStrRec)
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNewTStrRec As SAPCommon.TStrRec
        Dim aNewHdrRec As New TDataRec(aIntPar)
        Dim aStrucName() As String
        Dim aLen As Integer = 0
        For Each aKvb In aPar.getData()
            aTStrRec = aKvb.Value
            If valid_Import_Field(aTStrRec) Then
                aNewHdrRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pEmptyChar:="")
            End If
            aStrucName = Split(aTStrRec.Strucname, "+")
            For s As Integer = 0 To aStrucName.Length - 1
                aNewTStrRec = New SAPCommon.TStrRec(aStrucName(s), aTStrRec.Fieldname, aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
                If valid_Structure_Field(aNewTStrRec) Then
                    aNewHdrRec.setValues(aNewTStrRec.getKey(), aNewTStrRec.Value, aNewTStrRec.Currency, aNewTStrRec.Format, pEmptyChar:="")
                End If
            Next
        Next
        ' First fill the value from the paramters and tehn overwrite them from the posting record
        Dim aPostRec As New TDataRec(aIntPar)
        aPostRec = pData.getFirstRecord()
        For Each aTStrRec In aPostRec.aTDataRecCol
            If valid_Import_Field(aTStrRec) Then
                aNewHdrRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
            End If
            aStrucName = Split(aTStrRec.Strucname, "+")
            For s As Integer = 0 To aStrucName.Length - 1
                aNewTStrRec = New SAPCommon.TStrRec(aStrucName(s), aTStrRec.Fieldname, aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
                If valid_Structure_Field(aNewTStrRec) Then
                    aNewHdrRec.setValues(aNewTStrRec.getKey(), aNewTStrRec.Value, aNewTStrRec.Currency, aNewTStrRec.Format, pEmptyChar:="")
                End If
            Next
        Next
        aHdrRec = aNewHdrRec
        fillHeader = True
    End Function

    Public Function fillData(pData As TData) As Boolean
        Dim aKvB As KeyValuePair(Of String, TDataRec)
        Dim aTDataRec As TDataRec
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aCnt As UInt64
        Dim aStrucName() As String
        aDataDic = New TDataDic(aIntPar)
        fillData = True
        For Each aKvB In pData.aTDataDic
            aTDataRec = aKvB.Value
            For Each aTStrRec In aTDataRec.aTDataRecCol
                aStrucName = Split(aTStrRec.Strucname, "+")
                For s As Integer = 0 To aStrucName.Length - 1
                    If valid_Table_Field(aTStrRec, aStrucName(s)) Then
                        aDataDic.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=aStrucName(s))
                    End If
                Next
            Next
        Next

    End Function

    Public Function valid_Import_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        valid_Import_Field = False
        If pTStrRec.Strucname = "" Or pTStrRec.Strucname = "I" Then
            If aParamDic.ContainsKey("I|" & pTStrRec.Fieldname) Then
                valid_Import_Field = True
            End If
        End If
    End Function

    Public Function valid_Structure_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aStrucName() As String
        valid_Structure_Field = False
        aStrucName = Split(pTStrRec.Strucname, "+")
        For s As Integer = 0 To aStrucName.Length - 1
            If aStrucDic.ContainsKey("S|" & aStrucName(s)) Then
                valid_Structure_Field = isInStructure(pTStrRec.Fieldname, aStrucDic("S|" & aStrucName(s)))
            End If
        Next
    End Function

    Public Function valid_Table_Field(pTStrRec As SAPCommon.TStrRec, pStrucname As String) As Boolean
        valid_Table_Field = False
        If aStrucDic.ContainsKey("T|" & pStrucname) Then
            valid_Table_Field = isInStructure(pTStrRec.Fieldname, aStrucDic("T|" & pStrucname))
        End If
    End Function

    Private Function isInStructure(pName As String, pRfcStructureMetadata As RfcStructureMetadata, Optional ByRef pLen As Integer = 0) As Boolean
        Dim aRfcFieldMetadata As RfcFieldMetadata
        Try
            aRfcFieldMetadata = pRfcStructureMetadata.Item(pName)
            isInStructure = True
            pLen = aRfcFieldMetadata.NucLength
        Catch ex As Exception
            isInStructure = False
            pLen = 0
        End Try
    End Function

    Private Function isInArray(pString As String, pArray As Object) As Boolean
        Dim st As String, M As String
        M = "$"
        st = M & Join(pArray, M) & M
        isInArray = InStr(st, M & pString & M) > 0
        ' isInArray = (UBound(Filter(pArray, pString)) > -1)
    End Function

    Public Sub dumpHeader()
        Dim dumpHd As String = If(aIntPar.value("PROJ_DBG", "DUMPDATA") <> "", aIntPar.value("PROJ_DBG", "DUMPDATA"), "")
        If dumpHd <> "" Then
            Dim aDWS As Excel.Worksheet
            Dim aWB As Excel.Workbook
            Dim aRange As Excel.Range
            aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
            Try
                aDWS = aWB.Worksheets(dumpHd)
                aDWS.Activate()
            Catch Exc As System.Exception
                log.Warn("dumpWbsinfo - " & "No " & dumpHd & " Sheet in current workbook.")
                MsgBox("No " & dumpHd & " Sheet in current workbook. Check the PROJ_DBG-DUMPDATA Parameter",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
                Exit Sub
            End Try
            log.Debug("dumpHeader - " & "dumping to " & dumpHd)
            ' clear the Header
            If CStr(aDWS.Cells(1, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(1, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If
            ' dump the Header
            Dim aTStrRec As New SAPCommon.TStrRec
            Dim aFieldArray() As String = {}
            Dim aValueArray() As String = {}
            For Each aTStrRec In aHdrRec.aTDataRecCol
                Array.Resize(aFieldArray, aFieldArray.Length + 1)
                aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                Array.Resize(aValueArray, aValueArray.Length + 1)
                aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
            Next
            aRange = aDWS.Range(aDWS.Cells(1, 1), aDWS.Cells(1, aFieldArray.Length))
            aRange.Value = aFieldArray
            aRange = aDWS.Range(aDWS.Cells(2, 1), aDWS.Cells(2, aValueArray.Length))
            aRange.Value = aValueArray
        End If
    End Sub

    Public Sub dumpData()
        Dim dumpDt As String = If(aIntPar.value("PROJ_DBG", "DUMPDATA") <> "", aIntPar.value("PROJ_DBG", "DUMPDATA"), "")
        If dumpDt <> "" Then
            Dim aDWS As Excel.Worksheet
            Dim aWB As Excel.Workbook
            Dim aRange As Excel.Range
            aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
            Try
                aDWS = aWB.Worksheets(dumpDt)
                aDWS.Activate()
            Catch Exc As System.Exception
                log.Warn("dumpData - " & "No " & dumpDt & " Sheet in current workbook.")
                MsgBox("No " & dumpDt & " Sheet in current workbook. Check the PROJ_DBG-DUMPDATA Parameter",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Accounting")
                Exit Sub
            End Try
            log.Debug("dumpData - " & "dumping to " & dumpDt)
            ' clear the Data
            If CStr(aDWS.Cells(5, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(5, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If

            Dim aKvB_Dic As KeyValuePair(Of String, TData)
            Dim aKvB_Rec As KeyValuePair(Of String, TDataRec)
            Dim aData As TData
            Dim aData_Am As New TData(aIntPar)
            Dim aDataRec As New TDataRec(aIntPar)
            Dim aDataRec_Am As New TDataRec(aIntPar)
            Dim i As Int64
            Dim aTStrRec As New SAPCommon.TStrRec
            i = 6
            For Each aKvB_Dic In aDataDic.aTDataDic
                aData = aKvB_Dic.Value
                aDWS.Cells(i, 1).Value = aKvB_Dic.Key
                For Each aKvB_Rec In aData.aTDataDic
                    aDataRec = aKvB_Rec.Value
                    Dim aFieldArray() As String = {}
                    Dim aValueArray() As String = {}
                    For Each aTStrRec In aDataRec.aTDataRecCol
                        Array.Resize(aFieldArray, aFieldArray.Length + 1)
                        aFieldArray(aFieldArray.Length - 1) = aTStrRec.getKey()
                        Array.Resize(aValueArray, aValueArray.Length + 1)
                        aValueArray(aValueArray.Length - 1) = aTStrRec.formated()
                    Next
                    aRange = aDWS.Range(aDWS.Cells(i, 1), aDWS.Cells(i, aFieldArray.Length))
                    aRange.Value = aFieldArray
                    aRange = aDWS.Range(aDWS.Cells(i + 1, 1), aDWS.Cells(i + 1, aValueArray.Length))
                    aRange.Value = aValueArray
                    i += 2
                Next
                i += 2
            Next
        End If
    End Sub

End Class
