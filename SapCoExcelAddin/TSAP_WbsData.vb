' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class TSAP_WbsData

    Public aHdrRec As TDataRec
    Public aData As TData
    Public aExt As TData

    Private Hd_Fields() As String = {"I_PROJECT_DEFINITION"}
    Private Data_Fields() As String = {"WBS_ELEMENT", "DESCRIPTION", "RESPONSIBLE_NO", "APPLICANT_NO", "COMPANY_CODE", "BUSINESS_AREA", "CONTROLLING_AREA", "PROFIT_CTR", "PROJ_TYPE", "WBS_PLANNING_ELEMENT", "WBS_ACCOUNT_ASSIGNMENT_ELEMENT", "WBS_BILLING_ELEMENT", "CSTG_SHEET", "OVERHEAD_KEY", "RES_ANAL_KEY", "REQUEST_CCTR_CONTROLLING_AREA", "REQUEST_CCTR", "RESPSBL_CCTR_CONTROLLING_AREA", "RESPSBL_CCTR", "CALENDAR", "PRIORITY", "EQUIPMENT", "FUNCT_LOC", "CURRENCY", "CURRENCY_ISO", "PLANT", "USER_FIELD_KEY", "USER_FIELD_CHAR20_1", "USER_FIELD_CHAR20_2", "USER_FIELD_CHAR10_1", "USER_FIELD_CHAR10_2", "USER_FIELD_QUAN1", "USER_FIELD_UNIT1", "USER_FIELD_UNIT1_ISO", "USER_FIELD_QUAN2", "USER_FIELD_UNIT2", "USER_FIELD_UNIT2_ISO", "USER_FIELD_CURR1", "USER_FIELD_CUKY1", "USER_FIELD_CUKY1_ISO", "USER_FIELD_CURR2", "USER_FIELD_CUKY2", "USER_FIELD_CUKY2_ISO", "USER_FIELD_DATE1", "USER_FIELD_DATE2", "USER_FIELD_FLAG1", "USER_FIELD_FLAG2", "WBS_CCTR_POSTED_ACTUAL", "WBS_SUMMARIZATION", "OBJECTCLASS", "STATISTICAL", "TAXJURCODE", "INTEREST_PROF", "INVEST_PROFILE", "EVGEW", "CHANGE_NO", "SUBPROJECT", "PLANINTEGRATED", "INV_REASON", "SCALE", "ENVIR_INVEST", "REQUEST_COMP_CODE", "WBS_MRP_ELEMENT", "LOCATION", "VENTURE", "REC_IND", "EQUITY_TYP", "JV_OTYPE", "JV_JIBCL", "JV_JIBSA", "WBS_BASIC_START_DATE", "WBS_BASIC_FINISH_DATE", "WBS_FORECAST_START_DATE", "WBS_FORECAST_FINISH_DATE", "WBS_ACTUAL_START_DATE", "WBS_ACTUAL_FINISH_DATE", "WBS_BASIC_DURATION", "WBS_BASIC_DUR_UNIT", "WBS_BASIC_DUR_UNIT_ISO", "WBS_FORECAST_DURATION", "WBS_FORCAST_DUR_UNIT", "WBS_FORECAST_DUR_UNIT_ISO", "WBS_ACTUAL_DURATION", "WBS_ACTUAL_DUR_UNIT", "WBS_ACTUAL_DUR_UNIT_ISO", "WBS_LEFT", "WBS_UP", "FUNC_AREA"}

    Private aPar As SAPCommon.TStr
    Private aIntPar As SAPCommon.TStr
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Const sWbs As String = "IT_WBS_ELEMENT"

    Private aUseAsEmpty As String = "#"
    Public Sub New(ByRef pPar As SAPCommon.TStr, ByRef pIntPar As SAPCommon.TStr)
        aPar = pPar
        aIntPar = pIntPar
        aUseAsEmpty = If(aIntPar.value("GEN", "USEASEMPTY") <> "", aIntPar.value("GEN", "USEASEMPTY"), "#")
    End Sub

    Public Function fillHeader(pData As TData) As Boolean
        aHdrRec = New TDataRec(aIntPar)
        Dim aPostRec As New TDataRec(aIntPar)
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aNewHdrRec As New TDataRec(aIntPar)
        aPostRec = pData.getFirstRecord()
        If Not IsNothing(aPostRec) Then
            For Each aTStrRec In aPostRec.aTDataRecCol
                If valid_Hdr_Field(aTStrRec) Then
                    aNewHdrRec.setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pUseAsEmpty:=aUseAsEmpty)
                End If
            Next
        End If
        aHdrRec = aNewHdrRec
        fillHeader = True
    End Function

    Public Function fillData(pData As TData) As Boolean
        Dim aKvB As KeyValuePair(Of String, TDataRec)
        Dim aTDataRec As TDataRec
        Dim aTStrRec As SAPCommon.TStrRec
        Dim aWbsRec As SAPCommon.TStrRec
        Dim aCnt As UInt64
        aData = New TData(aIntPar)
        aExt = New TData(aIntPar)
        fillData = True
        aCnt = 1
        For Each aKvB In pData.aTDataDic
            aTDataRec = aKvB.Value
            aWbsRec = aTDataRec.getWbsRec
            ' add the valid WBS fields
            For Each aTStrRec In aTDataRec.aTDataRecCol
                If valid_Data_Field(aTStrRec) Then
                    aData.addValue(CStr(aCnt), aTStrRec, pNewStrucname:=sWbs, pUseAsEmpty:=aUseAsEmpty)
                ElseIf valid_Ext_Field(aTStrRec) Then
                    aExt.addValue(CStr(aCnt), aWbsRec, pUseAsEmpty:=aUseAsEmpty)
                    aExt.addValue(CStr(aCnt), aTStrRec, pUseAsEmpty:=aUseAsEmpty)
                End If
            Next
            aCnt += 1
        Next
    End Function

    Public Function valid_Hdr_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        valid_Hdr_Field = False
        If pTStrRec.Strucname = "" Or pTStrRec.Strucname = "HD" Then
            valid_Hdr_Field = isInArray(pTStrRec.Fieldname, Hd_Fields)
        End If
    End Function

    Public Function valid_Ext_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        Dim aValExtString As String = If(aIntPar.value("WBS_STR", "VALEXT") <> "", aIntPar.value("WBS_STR", "VALEXT"), "")
        valid_Ext_Field = False
        If pTStrRec.Strucname = aValExtString Then
            valid_Ext_Field = True
        End If
    End Function

    Public Function valid_Data_Field(pTStrRec As SAPCommon.TStrRec) As Boolean
        valid_Data_Field = False
        If pTStrRec.Strucname = "IT_WBS_ELEMENT" Or pTStrRec.Strucname = "WBS" Then
            valid_Data_Field = isInArray(pTStrRec.Fieldname, Data_Fields)
        End If
    End Function

    Private Function isInArray(pString As String, pArray As Object) As Boolean
        Dim st As String, M As String
        M = "$"
        st = M & Join(pArray, M) & M
        isInArray = InStr(st, M & pString & M) > 0
        ' isInArray = (UBound(Filter(pArray, pString)) > -1)
    End Function

    Public Function getProject() As String
        Dim aTStrRec As SAPCommon.TStrRec
        getProject = ""
        For Each aTStrRec In aHdrRec.aTDataRecCol
            If aTStrRec.Fieldname = "I_PROJECT_DEFINITION" Then
                getProject = aTStrRec.Value
                Exit Function
            End If
        Next
    End Function

    Public Sub dumpHeader()
        Dim dumpHd As String = If(aIntPar.value("WBS_DBG", "DUMPHEADER") <> "", aIntPar.value("WBS_DBG", "DUMPHEADER"), "")
        If dumpHd <> "" Then
            Dim aDWS As Excel.Worksheet
            Dim aWB As Excel.Workbook
            Dim aRange As Excel.Range
            aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
            Try
                aDWS = aWB.Worksheets(dumpHd)
                aDWS.Activate()
            Catch Exc As System.Exception
                log.Warn("dumpHeader - " & "No " & dumpHd & " Sheet in current workbook.")
                MsgBox("No " & dumpHd & " Sheet in current workbook. Check the WBS_DBG-DUMPHEADR Parameter",
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
        Dim dumpDt As String = If(aIntPar.value("WBS_DBG", "DUMPDATA") <> "", aIntPar.value("WBS_DBG", "DUMPDATA"), "")
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
                MsgBox("No " & dumpDt & " Sheet in current workbook. Check the WBS_DBG-DUMPDATA Parameter",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
                Exit Sub
            End Try
            log.Debug("dumpData - " & "dumping to " & dumpDt)
            ' clear the Data
            If CStr(aDWS.Cells(5, 1).Value) <> "" Then
                aRange = aDWS.Range(aDWS.Cells(5, 1), aDWS.Cells(1000, 1))
                aRange.EntireRow.Delete()
            End If

            Dim aKvB As KeyValuePair(Of String, TDataRec)
            Dim aData_Am As New TData(aIntPar)
            Dim aDataRec As New TDataRec(aIntPar)
            Dim aDataRec_Am As New TDataRec(aIntPar)
            Dim i As Int64
            Dim aTStrRec As New SAPCommon.TStrRec
            i = 6
            For Each aKvB In aData.aTDataDic
                aDataRec = aKvB.Value
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
        End If
    End Sub

End Class
