' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SapCoRibbonGenerate

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Function getIntParameters(ByRef pIntPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim i As Integer

        log.Debug("getIntParameters - " & "reading Parameter")
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter_Int")
        Catch Exc As System.Exception
            MsgBox("No Parameter_Int Sheet in current workbook. Check if the current workbook is a valid SAP PS Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO")
            getIntParameters = False
            Exit Function
        End Try
        i = 2
        pIntPar = New SAPCommon.TStr
        Do
            pIntPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
        ' no obligatory parameters check - we should know what we are doing
        getIntParameters = True
    End Function

    Public Sub exec(ByRef pSapCon As SapCon, Optional pNr As String = "")
        Dim aIntPar As New SAPCommon.TStr
        Dim aMigHelper As MigHelper
        Dim aBasis As New Collection
        Dim aBasisLine As New Dictionary(Of String, SAPCommon.TField)
        Dim aPostingLine As New Dictionary(Of String, SAPCommon.TField)
        Dim aContraLine As New Dictionary(Of String, SAPCommon.TField)
        Dim aPostingLines As Collection
        Dim aPWs As Excel.Worksheet
        Dim aIWs As Excel.Worksheet
        Dim aDWs As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aUseBasis As Boolean = False
        Dim i As Integer
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If
        ' get the ruleset limits
        Dim aDwsName As String = If(aIntPar.value("GEN" & pNr, "WS_DATA") <> "", aIntPar.value("GEN" & pNr, "WS_DATA"), "Data")
        Dim aBwsName As String = If(aIntPar.value("GEN" & pNr, "WS_BASE") <> "", aIntPar.value("GEN" & pNr, "WS_BASE"), "Basis")
        Dim aEmptyChar As String = If(aIntPar.value("GEN" & pNr, "CHAR_EMPTY") <> "", aIntPar.value("GEN" & pNr, "CHAR_EMPTY"), "#")
        Dim aIgnoreEmpty As String = If(aIntPar.value("GEN" & pNr, "IGNORE_EMPTY") <> "", aIntPar.value("GEN" & pNr, "IGNORE_EMPTY"), "X")
        Dim aGenEmpty As Boolean = If(aIgnoreEmpty = "X", False, True)
        Dim aDeleteData As String = If(aIntPar.value("GEN" & pNr, "DELETE_DATA") <> "", aIntPar.value("GEN" & pNr, "DELETE_DATA"), "X")
        Dim aGenDeleteData As Boolean = If(aDeleteData = "X", True, False)
        Dim aLOff As Integer = If(aIntPar.value("GEN" & pNr, "LOFF_DATA") <> "", CInt(aIntPar.value("GEN" & pNr, "LOFF_DATA")), 4)
        Dim aLOffBData As Integer = If(aIntPar.value("GEN" & pNr, "LOFF_BDATA") <> "", CInt(aIntPar.value("GEN" & pNr, "LOFF_BDATA")), 1)
        Dim aLOffBNames As Integer = If(aIntPar.value("GEN" & pNr, "LOFF_BNAMES") <> "", CInt(aIntPar.value("GEN" & pNr, "LOFF_BNAMES")), 0)
        Dim aLOffTNames As Integer = If(aIntPar.value("GEN" & pNr, "LOFF_TNAMES") <> "", CInt(aIntPar.value("GEN" & pNr, "LOFF_TNAMES")), aLOff - 1)
        Dim aLineOut As Integer = If(aIntPar.value("GEN" & pNr, "LINE_OUT") <> "", CInt(aIntPar.value("GEN" & pNr, "LINE_OUT")), 0)
        Dim aBaseColFrom As Integer = If(aIntPar.value("GEN" & pNr, "BASE_COLFROM") <> "", CInt(aIntPar.value("GEN" & pNr, "BASE_COLFROM")), 1)
        Dim aBaseColTo As Integer = If(aIntPar.value("GEN" & pNr, "BASE_COLTO") <> "", CInt(aIntPar.value("GEN" & pNr, "BASE_COLTO")), 100)
        Dim aBaseFilter As String = If(aIntPar.value("GEN" & pNr, "BASE_FILTER") <> "", CStr(aIntPar.value("GEN" & pNr, "BASE_FILTER")), "")
        Dim aTargetFilter As String = If(aIntPar.value("GEN" & pNr, "TARGET_FILTER") <> "", CStr(aIntPar.value("GEN" & pNr, "TARGET_FILTER")), "")
        ' should we compress posting lines?
        Dim aGenCompData As String = If(aIntPar.value("GEN" & pNr, "COMP_DATA") <> "", CStr(aIntPar.value("GEN" & pNr, "COMP_DATA")), "")
        Dim aCompress As Boolean = If(aGenCompData = "X", True, False)
        ' should we suppress line with zero values
        Dim aGenSupprZero As String = If(aIntPar.value("GEN" & pNr, "SUPPR_ZERO") <> "", CStr(aIntPar.value("GEN" & pNr, "SUPPR_ZERO")), "")
        Dim aSupprZero As Boolean = If(aGenSupprZero = "X", True, False)
        log.Debug("SapCoRibbonGenerate.exec - " & "Basis Sheet")
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aIWs = aWB.Worksheets(aBwsName)
            aUseBasis = True
        Catch Exc As System.Exception
            log.Warn("SapCoRibbonGenerate.exec - " & "No InvoiceData or " & aBwsName & " in current workbook.")
            MsgBox("No " & aBwsName & " in current workbook. Check if the current workbook is a valid Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO")
            Exit Sub
        End Try
        aMigHelper = New MigHelper(aIntPar, pNr, pBaseFilterStr:=aBaseFilter)
        ' process the data
        Try
            log.Debug("SapCoRibbonGenerate.exec - " & "processing data - disabling events, screen update, cursor")
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapCoExcelAddin.Application.EnableEvents = False
            Globals.SapCoExcelAddin.Application.ScreenUpdating = False
            ' read the invoice reposting lines
            i = aLOffBData + 1
            Do
                If aUseBasis Then
                    aBasisLine = aMigHelper.makeDictForRules(aIWs, i, aLOffBNames + 1, aBaseColFrom, aBaseColTo)
                    If Not aMigHelper.isFiltered(aBasisLine) Then
                        aBasis.Add(aBasisLine)
                    End If
                End If
                i = i + 1
            Loop While CStr(aIWs.Cells(i, 1).value) <> ""

            Dim aTPostingData As New TPostingData(aIntPar)
            Dim aTPostingDataRec As TPostingDataRec
            Dim aTPostingDataRecKey As String
            Dim aTPostingDataRecNum As UInt64 = 1
            ' create the posting lines
            aPostingLines = New Collection
            For Each aBasisLine In aBasis
                aPostingLine = aMigHelper.mig.ApplyRules(aBasisLine, "P")
                If aPostingLine.Count > 0 Then
                    aTPostingDataRec = aTPostingData.newTPostingDataRec(pDic:=aPostingLine, pEmpty:=aGenEmpty, pEmptyChar:=aEmptyChar, pNr:=pNr)
                    If aCompress Then
                        aTPostingDataRecKey = aTPostingDataRec.getKey()
                    Else
                        aTPostingDataRecKey = CStr(aTPostingDataRecNum)
                    End If
                    aTPostingData.addTPostingDataRec(aTPostingDataRecKey, aTPostingDataRec, pNr:=pNr)
                    aTPostingDataRecNum += 1
                End If
                aContraLine = aMigHelper.mig.ApplyRules(aBasisLine, "C")
                If aContraLine.Count > 0 Then
                    aTPostingDataRec = aTPostingData.newTPostingDataRec(pDic:=aContraLine, pEmpty:=aGenEmpty, pEmptyChar:=aEmptyChar, pNr:=pNr)
                    If aCompress Then
                        aTPostingDataRecKey = aTPostingDataRec.getKey()
                    Else
                        aTPostingDataRecKey = CStr(aTPostingDataRecNum)
                    End If
                    aTPostingData.addTPostingDataRec(aTPostingDataRecKey, aTPostingDataRec, pNr:=pNr)
                    aTPostingDataRecNum += 1
                End If
            Next
            'output the posting lines
            Dim aColDelData As Integer = If(aIntPar.value("GEN" & pNr, "DATA_COLDEL") <> "", CInt(aIntPar.value("GEN" & pNr, "DATA_COLDEL")), 1)
            Try
                aDWs = aWB.Worksheets(aDwsName)
            Catch Exc As System.Exception
                Globals.SapCoExcelAddin.Application.EnableEvents = True
                Globals.SapCoExcelAddin.Application.ScreenUpdating = True
                Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
                MsgBox("No " & aDwsName & " Sheet in current workbook. Check if the current workbook is a valid Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO")
                Exit Sub
            End Try
            Dim aRange As Excel.Range
            i = aLOff + 1
            Do
                i += 1
            Loop While CStr(aDWs.Cells(i, aColDelData).Value) <> ""
            If aGenDeleteData And i >= aLOff + 1 Then
                aRange = aDWs.Range(aDWs.Cells(aLOff + 1, 1), aDWs.Cells(i, 1))
                aRange.EntireRow.Delete()
                i = aLOff + 1
            End If
            Dim jMax As Integer = 0
            Do
                jMax += 1
            Loop While CStr(aDWs.Cells(aLOff, jMax + 1).value) <> ""
            Dim aKey As String
            Dim aValue As String
            Dim aKvb As KeyValuePair(Of String, TPostingDataRec)
            i = If(aLineOut <> 0, aLineOut, i)
            Dim aSuppressLine As Boolean
            For Each aKvb In aTPostingData.aTPostingDataDic
                aSuppressLine = False
                aTPostingDataRec = aKvb.Value
                If aSupprZero And aTPostingDataRec.isZero Then
                    aSuppressLine = True
                End If
                If Not String.IsNullOrEmpty(aTargetFilter) Then
                    If isTargetFiltered(aTargetFilter, aTPostingDataRec) Then
                        aSuppressLine = True
                    End If
                End If
                    If Not aSuppressLine Then
                    For j = 1 To jMax
                        If CStr(aDWs.Cells(aLOffTNames + 1, j).Value) <> "" Then
                            aKey = CStr(aDWs.Cells(aLOffTNames + 1, j).Value)
                            If Not aKey.Contains("-") Then
                                aKey = "-" & aKey
                            End If
                            If aTPostingDataRec.aTPostingDataRecCol.Contains(aKey) Then
                                aValue = aTPostingDataRec.aTPostingDataRecCol(aKey).Value
                                If aTPostingDataRec.isValue(aKey) Then
                                    aDWs.Cells(i, j).Value = CDbl(aValue)
                                ElseIf aTPostingDataRec.aTPostingDataRecCol(aKey).Format = "F" Then
                                    aDWs.Cells(i, j).FormulaR1C1 = "=" & CStr(aValue)
                                Else
                                    aDWs.Cells(i, j).Value = aValue
                                End If
                            End If
                        End If
                    Next j
                    i += 1
                End If
            Next
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            aDWs.Activate()
        Catch ex As System.Exception
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapCoRibbonGenerate.exec failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO")
            log.Error("SapCoRibbonGenerate.exec - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Private Function isTargetFiltered(pTargetFilterStr As String, pTPostingDataRec As TPostingDataRec) As Boolean
        Dim aFilterField As String = ""
        Dim aFilterOperation As String = ""
        Dim aFilterCompare As String = ""
        If Not String.IsNullOrEmpty(pTargetFilterStr) Then
            Dim aFilterStr() As String = {}
            aFilterStr = pTargetFilterStr.Split(";")
            If aFilterStr.Length = 3 Then
                aFilterField = aFilterStr(0)
                aFilterOperation = aFilterStr(1)
                aFilterCompare = aFilterStr(2)
                If aFilterCompare.ToUpper() = "NULL" Then
                    aFilterCompare = ""
                End If
            End If
        End If
        isTargetFiltered = False
        Dim aTStrRec As SAPCommon.TStrRec
        If pTPostingDataRec.aTPostingDataRecCol.Contains("-" & aFilterField) Then
            aTStrRec = pTPostingDataRec.aTPostingDataRecCol("-" & aFilterField)
            If aFilterOperation = "EQ" And aTStrRec.Value = aFilterCompare Then
                isTargetFiltered = True
            ElseIf aFilterOperation = "NE" And aTStrRec.Value <> aFilterCompare Then
                isTargetFiltered = True
            End If
        Else
            If aFilterOperation = "NE" And (String.IsNullOrEmpty(aFilterCompare) Or aFilterCompare = "#") Then
                isTargetFiltered = False
            ElseIf aFilterOperation = "EQ" And (String.IsNullOrEmpty(aFilterCompare) Or aFilterCompare = "#") Then
                isTargetFiltered = True
            End If
        End If
    End Function

    Public Sub execWbs(ByRef pSapCon As SapCon)
        Dim aIntPar As New SAPCommon.TStr
        Dim aWBSTemplate As Collection = New Collection
        Dim aWBSList As SortedList(Of String, String()) = New SortedList(Of String, String())
        Dim aWWs As Excel.Worksheet
        Dim aBWs As Excel.Worksheet
        Dim aDWs As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim i As Integer
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If
        ' get the ruleset limits
        Dim aDwsName As String = If(aIntPar.value("GENWBS", "WS_DATA") <> "", aIntPar.value("GENWBS", "WS_DATA"), "Data")
        Dim aBwsName As String = If(aIntPar.value("GENWBS", "WS_BASE") <> "", aIntPar.value("GENWBS", "WS_BASE"), "Basis")
        Dim aWwsName As String = If(aIntPar.value("GENWBS", "WS_WBS") <> "", aIntPar.value("GENWBS", "WS_WBS"), "WBS_Read")
        Dim aDeleteData As String = If(aIntPar.value("GENWBS", "DELETE_DATA") <> "", aIntPar.value("GENWBS", "DELETE_DATA"), "X")
        Dim aGenDeleteData As Boolean = If(aDeleteData = "X", True, False)
        Dim aLOffBData As Integer = If(aIntPar.value("GENWBS", "LOFF_BDATA") <> "", CInt(aIntPar.value("GENWBS", "LOFF_BDATA")), 0)
        Dim aLOffWData As Integer = If(aIntPar.value("GENWBS", "LOFF_WDATA") <> "", CInt(aIntPar.value("GENWBS", "LOFF_WDATA")), 0)
        Dim aLOffData As Integer = If(aIntPar.value("GENWBS", "LOFF_DATA") <> "", CInt(aIntPar.value("GENWBS", "LOFF_DATA")), 0)
        Dim aBColWbs As Integer = If(aIntPar.value("GENWBS", "COL_BWBS") <> "", CInt(aIntPar.value("GENWBS", "COL_BWBS")), 1)
        Dim aWColWbs As Integer = If(aIntPar.value("GENWBS", "COL_WWBS") <> "", CInt(aIntPar.value("GENWBS", "COL_WWBS")), 1)
        Dim aWColUp As Integer = If(aIntPar.value("GENWBS", "COL_WUP") <> "", CInt(aIntPar.value("GENWBS", "COL_WUP")), 2)
        Dim aDColWbs As Integer = If(aIntPar.value("GENWBS", "COL_DWBS") <> "", CInt(aIntPar.value("GENWBS", "COL_DWBS")), 1)
        Dim aDColWbsT As Integer = If(aIntPar.value("GENWBS", "COL_DWBS_T") <> "", CInt(aIntPar.value("GENWBS", "COL_DWBS_T")), 2)
        Dim aDColWbsUpT As Integer = If(aIntPar.value("GENWBS", "COL_DUP_T") <> "", CInt(aIntPar.value("GENWBS", "COL_DUP_T")), 3)
        Dim aWPattern_New As String = If(aIntPar.value("GENWBS", "PATTERN_NEW") <> "", CStr(aIntPar.value("GENWBS", "PATTERN_NEW")), "")
        Dim aWPattern_Old As String = If(aIntPar.value("GENWBS", "PATTERN_OLD") <> "", CStr(aIntPar.value("GENWBS", "PATTERN_OLD")), "")
        log.Debug("SapCoRibbonGenerate.execWbs - " & "WBS Sheet")
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aWWs = aWB.Worksheets(aWwsName)
        Catch Exc As System.Exception
            log.Warn("SapCoRibbonGenerate.execWbs - " & "No InvoiceData or " & aWwsName & " in current workbook.")
            MsgBox("No " & aWwsName & " in current workbook. Check if the current workbook is a valid Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO")
            Exit Sub
        End Try
        log.Debug("SapCoRibbonGenerate.execWbs - " & "Basis Sheet")
        Try
            aBWs = aWB.Worksheets(aBwsName)
        Catch Exc As System.Exception
            log.Warn("SapCoRibbonGenerate.execWbs - " & "No InvoiceData or " & aBwsName & " in current workbook.")
            MsgBox("No " & aBwsName & " in current workbook. Check if the current workbook is a valid Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO")
            Exit Sub
        End Try
        Try
            log.Debug("SapCoRibbonGenerate.execWbs - " & "processing data - disabling events, screen update, cursor")
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapCoExcelAddin.Application.EnableEvents = False
            Globals.SapCoExcelAddin.Application.ScreenUpdating = False
            ' build the WBS collection
            i = aLOffWData + 1
            Dim aMaxJ As UInt64 = 0
            Do
                aWBSTemplate.Add(CStr(aWWs.Cells(i, aWColUp).Value), CStr(aWWs.Cells(i, aWColWbs).Value))
                i += 1
            Loop While CStr(aWWs.Cells(i, 1).value) <> ""
            ' process the base data
            i = aLOffBData + 1
            Do
                getWBSwithParents(CStr(aBWs.Cells(i, aBColWbs).Value), aWBSTemplate, aWBSList, aWPattern_New, aWPattern_Old)
                i += 1
            Loop While CStr(aBWs.Cells(i, 1).value) <> ""
            'output the data
            Dim aColDelData As Integer = If(aIntPar.value("GENWBS", "DATA_COLDEL") <> "", CInt(aIntPar.value("GENWBS", "DATA_COLDEL")), 1)
            Try
                aDWs = aWB.Worksheets(aDwsName)
            Catch Exc As System.Exception
                Globals.SapCoExcelAddin.Application.EnableEvents = True
                Globals.SapCoExcelAddin.Application.ScreenUpdating = True
                Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
                MsgBox("No " & aDwsName & " Sheet in current workbook. Check if the current workbook is a valid Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO")
                Exit Sub
            End Try
            Dim aRange As Excel.Range
            i = aLOffData + 1
            Do
                i += 1
            Loop While CStr(aDWs.Cells(i, aColDelData).Value) <> ""
            If aGenDeleteData And i >= aLOffData + 1 Then
                aRange = aDWs.Range(aDWs.Cells(aLOffData + 1, 1), aDWs.Cells(i, 1))
                aRange.EntireRow.Delete()
                i = aLOffData + 1
            End If
            Dim aWbsKey As String = ""
            For Each aWbsKey In aWBSList.Keys
                aDWs.Cells(i, aDColWbs).Value = aWbsKey
                Dim aWbsArray = aWBSList(aWbsKey)
                aDWs.Cells(i, aDColWbsT).Value = aWbsArray(0)
                aDWs.Cells(i, aDColWbsUpT).Value = aWbsArray(1)
                i += 1
            Next
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            aDWs.Activate()
        Catch ex As System.Exception
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapCoRibbonGenerate.execWbs failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO")
            log.Error("SapCoRibbonGenerate.execWbs - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Private Sub getWBSwithParents(pWBS As String, ByRef pWBSTemplate As Collection, pWBSList As SortedList(Of String, String()), pPattNew As String, pPattOld As String)
        Dim aSapFormat As New SAPCommon.SAPFormat
        If pWBSList.Keys.Contains(pWBS) Or pWBSTemplate.Contains(pWBS) Then
            Exit Sub
        Else
            Dim aWbsArray(2) As String
            aWbsArray(0) = SAPCommon.SAPFormat.applyPattern(pWBS, pPattOld)
            aWbsArray(1) = pWBSTemplate(aWbsArray(0))
            pWBSList.Add(pWBS, aWbsArray)
            Dim aWbsOld As String = SAPCommon.SAPFormat.applyPattern(pWBS, pPattOld)
            If pWBSTemplate.Contains(aWbsOld) Then
                Dim aWbsUpOld = pWBSTemplate(aWbsOld)
                Dim aWbsUpNew = SAPCommon.SAPFormat.applyPattern(aWbsUpOld, pPattNew)
                getWBSwithParents(aWbsUpNew, pWBSTemplate, pWBSList, pPattNew, pPattOld)
            Else
                Exit Sub
            End If
        End If
    End Sub

End Class
