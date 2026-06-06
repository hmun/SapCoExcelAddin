' Copyright 2022-2025 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports System.Runtime.InteropServices.ComTypes
Imports System.Security.Cryptography
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Tools.Ribbon
Imports SAP.Middleware.Connector
Imports SAPCommon
Imports SAPLogon
Imports SC = SAPCommon

Public Class Ribbon_Generate

    Private app = Globals.ThisAddIn.Application
    Private aSapCon
    Private aSapGeneral

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Sub New(ByRef pSapGeneral As SapGeneral)
        aSapGeneral = pSapGeneral
    End Sub

    Private Function checkCon() As Integer
        Dim aSapConRet As Integer
        Dim aSapVersionRet As Integer
        checkCon = False
        log.Debug("checkCon - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            Exit Function
        End If
        log.Debug("checkCon - " & "checking Connection")
        aSapConRet = 0
        If aSapCon Is Nothing Then
            Try
                aSapCon = New SapCon()
            Catch ex As SystemException
                log.Warn("checkCon-New SapCon - )" & ex.ToString)
            End Try
        End If
        Try
            aSapConRet = aSapCon.checkCon()
        Catch ex As SystemException
            log.Warn("checkCon-aSapCon.checkCon - )" & ex.ToString)
        End Try
        If aSapConRet = 0 Then
            log.Debug("checkCon - " & "checking version in SAP")
            Try
                aSapVersionRet = aSapGeneral.checkVersionInSAP(aSapCon)
            Catch ex As SystemException
                log.Warn("checkCon - )" & ex.ToString)
            End Try
            log.Debug("checkCon - " & "aSapVersionRet=" & CStr(aSapVersionRet))
            If aSapVersionRet = True Then
                log.Debug("checkCon - " & "checkCon = True")
                checkCon = True
            Else
                log.Debug("checkCon - " & "connection check failed")
            End If
        End If
    End Function

    Public Function getIntParameters(ByRef pIntPar As SC.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim i As Integer

        log.Debug("getIntParameters - " & "reading Parameter")
        aWB = app.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter_Int")
        Catch Exc As System.Exception
            MsgBox("No Parameter_Int Sheet in current workbook. Check if the current workbook is a valid Migration Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Migration")
            getIntParameters = False
            Exit Function
        End Try
        i = 2
        pIntPar = New SC.TStr
        Do
            pIntPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
        ' no obligatory parameters check - we should know what we are doing
        getIntParameters = True
    End Function

    Public Sub GenerateData()
        Dim aIntPar As New SC.TStr
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If
        Try
            ' get the ruleset limits
            Dim aGenNrFrom As Integer = If(aIntPar.value("GEN", "RULESET_FROM") <> "", CInt(aIntPar.value("GEN", "RULESET_FROM")), 0)
            Dim aGenNrTo As Integer = If(aIntPar.value("GEN", "RULESET_TO") <> "", CInt(aIntPar.value("GEN", "RULESET_TO")), 0)
            Dim aGenNr As String = ""
            For i As Integer = aGenNrFrom To aGenNrTo
                Dim aNr As String = If(i = 0, "", CStr(i))
                Dim aRunBefore As String = If(aIntPar.value("GEN" & aNr, "RUN_BEFORE") <> "", aIntPar.value("GEN" & aNr, "RUN_BEFORE"), "")
                If aRunBefore = "GENWBS" Then
                    ' read the Template WBS Information from SAP
                    Dim aSapPsMdRibbonWbs As New SapPsMdRibbonWbs
                    If checkCon() = True Then
                        aSapPsMdRibbonWbs.GetData(pSapCon:=aSapCon)
                    Else
                        MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonGetWbs_Click")
                    End If
                    ' genertate the WBS-Elements
                    execWbs(pSapCon:=aSapCon)
                End If
                GenerateData_exec(pIntPar:=aIntPar, pNr:=aNr)
            Next
            Dim aSortNrFrom As Integer = If(aIntPar.value("SORT", "RULESET_FROM") <> "", CInt(aIntPar.value("SORT", "RULESET_FROM")), 0)
            Dim aSortNrTo As Integer = If(aIntPar.value("SORT", "RULESET_TO") <> "", CInt(aIntPar.value("SORT", "RULESET_TO")), 0)
            Dim aSortNr As String = ""
            For i As Integer = aSortNrFrom To aSortNrTo
                Dim aNr As String = If(i = 0, "", CStr(i))
                SortData_exec(pIntPar:=aIntPar, pNr:=aNr)
            Next
            Dim aDoubNrFrom As Integer = If(aIntPar.value("DOUB", "RULESET_FROM") <> "", CInt(aIntPar.value("DOUB", "RULESET_FROM")), 0)
            Dim aDoubNrTo As Integer = If(aIntPar.value("DOUB", "RULESET_TO") <> "", CInt(aIntPar.value("DOUB", "RULESET_TO")), 0)
            Dim aDoubNr As String = ""
            For i As Integer = aDoubNrFrom To aDoubNrTo
                Dim aNr As String = If(i = 0, "", CStr(i))
                RemoveDoulicates_exec(pIntPar:=aIntPar, pNr:=aNr)
            Next
            MsgBox("Generate completed", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Generation")
        Catch ex As System.Exception
            MsgBox("GenerateData failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Migration")
            log.Error("GenerateData - " & "Exception=" & ex.ToString)
        End Try
        app.EnableEvents = True
        app.ScreenUpdating = True
        app.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
    End Sub

    Private Sub GenerateData_exec(ByRef pIntPar As SC.TStr, Optional pNr As String = "")
        Dim aMigHelperVsto As MigHelperVsto
        Dim aBWs As Excel.Worksheet
        Dim aOWs As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aUselocal As Boolean = False
        Dim i As UInt32
        ' get internal parameters
        Dim aOwsName As String = If(pIntPar.value("GEN" & pNr, "WS_DATA") <> "", pIntPar.value("GEN" & pNr, "WS_DATA"), "Data")
        Dim aBwsName As String = If(pIntPar.value("GEN" & pNr, "WS_BASE") <> "", pIntPar.value("GEN" & pNr, "WS_BASE"), "Base")
        Dim aDeleteData As String = If(pIntPar.value("GEN" & pNr, "DELETE_DATA") <> "", pIntPar.value("GEN" & pNr, "DELETE_DATA"), "X")
        Dim aGenDeleteData As Boolean = If(aDeleteData = "X", True, False)
        Dim aLOff As Integer = If(pIntPar.value("GEN" & pNr, "LOFF_DATA") <> "", CInt(pIntPar.value("GEN" & pNr, "LOFF_DATA")), 4)
        Dim aLOffBData As Integer = If(pIntPar.value("GEN" & pNr, "LOFF_BDATA") <> "", CInt(pIntPar.value("GEN" & pNr, "LOFF_BDATA")), 1)
        Dim aLOffBNames As Integer = If(pIntPar.value("GEN" & pNr, "LOFF_BNAMES") <> "", CInt(pIntPar.value("GEN" & pNr, "LOFF_BNAMES")), 0)
        Dim aLOffTNames As Integer = If(pIntPar.value("GEN" & pNr, "LOFF_TNAMES") <> "", CInt(pIntPar.value("GEN" & pNr, "LOFF_TNAMES")), aLOff - 1)
        Dim aLineOut As Integer = If(pIntPar.value("GEN" & pNr, "LINE_OUT") <> "", CInt(pIntPar.value("GEN" & pNr, "LINE_OUT")), 0)
        Dim aBaseColFrom As Integer = If(pIntPar.value("GEN" & pNr, "BASE_COLFROM") <> "", CInt(pIntPar.value("GEN" & pNr, "BASE_COLFROM")), 1)
        Dim aBaseColTo As Integer = If(pIntPar.value("GEN" & pNr, "BASE_COLTO") <> "", CInt(pIntPar.value("GEN" & pNr, "BASE_COLTO")), 100)
        log.Debug("GenerateData_exec - " & "Basis Sheet")
        aWB = app.ActiveWorkbook
        Dim aGenLocalRules As String = If(pIntPar.value("GEN", "LOCAL_RULES") <> "", CStr(pIntPar.value("GEN", "LOCAL_RULES")), "")
        If aGenLocalRules = "X" Then
            aUselocal = True
            log.Debug("GenerateData_exec - " & "aUselocal = True")
        Else
            ' Fallback for compatibilty to old templates
            Try
                Dim aPWs As Excel.Worksheet = aWB.Worksheets("Parameter")
                If CStr(aPWs.Cells(17, 2).Value) = "X" Then
                    aUselocal = True
                    log.Debug("ButtonGenGLData_Click - " & "aUselocal = True")
                End If
            Catch Exc As System.Exception
                log.Debug("ButtonGenGLData_Click - " & "No Parameter Sheet in current workbook. -> aUselocal = False")
            End Try
        End If
        Try
            aBWs = aWB.Worksheets("InvoiceData")
        Catch Ex As System.Exception
            Try
                aBWs = aWB.Worksheets(aBwsName)
            Catch Exc As System.Exception
                log.Warn("GenerateData_exec - " & "No InvoiceData or " & aBwsName & " in current workbook.")
                MsgBox("No InvoiceData Sheet or " & aBwsName & " in current workbook. Check if the current workbook is a valid Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Migration")
                Exit Sub
            End Try
        End Try
        If String.IsNullOrEmpty(CStr(aBWs.Cells(aLOffBData + 1, aBaseColFrom).Value)) Then
            MsgBox("Base data cell row=" & aLOffBData + 1 & ", column=" & aBaseColFrom & " is empty. Check if the current workbook contains data and your internal parameters are correct!",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Migration")
            Exit Sub
        End If
        If String.IsNullOrEmpty(CStr(aBWs.Cells(aLOffBNames + 1, aBaseColFrom).Value)) Then
            MsgBox("Base data name cell row=" & aLOffBNames + 1 & ", column=" & aBaseColFrom & " is empty. Check if the current workbook contains data and your internal parameters are correct!",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Migration")
            Exit Sub
        End If
        '        aBWs.Activate()
        aMigHelperVsto = New MigHelperVsto(pIntPar:=pIntPar, pNr:=pNr, pUselocal:=aUselocal)
        ' process the data
        Try
            log.Debug("ButtonGenGLData_Click - " & "processing data - disabling events, screen update, cursor")
            app.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            app.EnableEvents = False
            app.ScreenUpdating = False
            ' read the base lines
            app.StatusBar = "Reading the base data"
            i = aLOffBData + 1
            Dim aLastRow As UInt64
            If Not String.IsNullOrEmpty(CStr(aBWs.Cells(aLOffBData + 2, aBaseColFrom).Value)) Then
                aLastRow = aBWs.Cells(aLOffBData, aBaseColFrom).End(Microsoft.Office.Interop.Excel.XlDirection.xlDown).Row
            Else
                aLastRow = aLOffBData + 1
            End If
            Dim aNamRange As Excel.Range = aBWs.Range(aBWs.Cells(aLOffBNames + 1, aBaseColFrom), aBWs.Cells(aLOffBNames + 1, aBaseColTo))
            Dim aNamArray As Object(,) = CType(aNamRange.Value, Object(,))
            Dim aValRange As Excel.Range = aBWs.Range(aBWs.Cells(aLOffBData + 1, aBaseColFrom), aBWs.Cells(aLastRow, aBaseColTo))
            Dim aValArray As Object(,) = CType(aValRange.Value, Object(,))
            Dim aMigEngine As SC.MigEngine = New SC.MigEngine(aMigHelperVsto.mh, pIntPar, pNr)
            ' migrating data
            app.StatusBar = "Migrating Data"
            aMigEngine.migrate(aNamArray, aValArray)
            ' prepare the output
            app.StatusBar = "Preparing Output"
            Dim aColDelData As Integer = If(pIntPar.value("SORT", "DATA_COLDEL") <> "", CInt(pIntPar.value("SORT", "DATA_COLDEL")), 1)
            Try
                aOWs = aWB.Worksheets(aOwsName)
            Catch Exc As System.Exception
                app.EnableEvents = True
                app.ScreenUpdating = True
                app.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
                MsgBox("No " & aOwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Migration Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Migration")
                Exit Sub
            End Try
            i = aLOff + 1
            aLastRow = aOWs.Cells(aOWs.Cells.Rows.Count, aColDelData).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            If aGenDeleteData And aLastRow >= aLOff + 1 Then
                app.StatusBar = "Deleting existing " & aLastRow - (aLOff + 1) & " lines in " & aOwsName
                Dim aRange As Excel.Range = aOWs.Range(aOWs.Cells(aLOff + 1, 1), aOWs.Cells(aLastRow, 1))
                Dim unused = aRange.EntireRow.Delete()
            End If
            Dim jMax As Integer = 0
            Do
                jMax += 1
            Loop While CStr(aOWs.Cells(aLOffTNames + 1, jMax + 1).value) <> ""
            Dim aOutLine = If(aLineOut <> 0, aLineOut, If(aGenDeleteData, aLOff + 1, aLastRow + 1))
            Dim aKeyRange As Excel.Range = aOWs.Range(aOWs.Cells(aLOff, 1), aOWs.Cells(aLOffTNames + 1, jMax))
            Dim aKeyArray As Object(,) = CType(aKeyRange.Value, Object(,))
            Dim aValueColumns As New SC.TOutData
            Dim aFormulaColumns As New SC.TOutData
            ' convert to output columns
            app.StatusBar = "Converting Output Columns"
            aMigEngine.ToTOutData(aKeyArray, aValueColumns, aFormulaColumns)
            ' write output to target sheet
            app.StatusBar = "Writing Value Columns"
            aMigHelperVsto.writeOutData(aOWs, aValueColumns, aOutLine, aKeyArray, pType:="V")
            app.StatusBar = "Writing Formula Columns"
            aMigHelperVsto.writeOutData(aOWs, aFormulaColumns, aOutLine, aKeyArray, pType:="F")
            aValueColumns = Nothing
            aFormulaColumns = Nothing
            aMigEngine = Nothing
            aMigHelperVsto = Nothing
            app.StatusBar = "Migration completed"
            app.EnableEvents = True
            app.ScreenUpdating = True
            app.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            app.EnableEvents = True
            app.ScreenUpdating = True
            app.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("ButtonGenGLData_Click failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Migration")
            log.Error("ButtonGenGLData_Click - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Private Sub SortData_exec(ByRef pIntPar As SC.TStr, Optional pNr As String = "")
        Dim aWB As Excel.Workbook
        Dim aOWs As Excel.Worksheet
        Dim aOwsName As String = If(pIntPar.value("SORT" & pNr, "WS_DATA") <> "", pIntPar.value("SORT" & pNr, "WS_DATA"), "Data")
        Dim aLOff As Integer = If(pIntPar.value("SORT" & pNr, "LOFF_DATA") <> "", CInt(pIntPar.value("SORT" & pNr, "LOFF_DATA")), 4)
        Dim aLOffTNames As Integer = If(pIntPar.value("SORT" & pNr, "LOFF_TNAMES") <> "", CInt(pIntPar.value("SORT" & pNr, "LOFF_TNAMES")), aLOff - 1)
        Dim aColDelData As Integer = If(pIntPar.value("SORT" & pNr, "DATA_CHEKCOL") <> "", CInt(pIntPar.value("SORT" & pNr, "DATA_CHEKCOL")), 1)
        Dim aColFrom As Integer = If(pIntPar.value("SORT" & pNr, "DATA_COLFROM") <> "", CInt(pIntPar.value("SORT" & pNr, "DATA_COLFROM")), 1)
        Dim aColTo As Integer = If(pIntPar.value("SORT" & pNr, "DATA_COLTO") <> "", CInt(pIntPar.value("SORT" & pNr, "DATA_COLTO")), 100)
        Dim aSortString As String = If(pIntPar.value("SORT" & pNr, "FIELDS") <> "", pIntPar.value("SORT" & pNr, "FIELDS"), "")
        Dim i As UInt32
        Dim aLastRow As UInt64
        Dim aNameDict As New Dictionary(Of String, UInt64)
        If aSortString = "" Then
            log.Debug("SortData_exec - " & "empty sort string - exit")
            Exit Sub
        End If
        aWB = app.ActiveWorkbook
        Try
            aOWs = aWB.Worksheets(aOwsName)
        Catch Exc As System.Exception
            app.EnableEvents = True
            app.ScreenUpdating = True
            app.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("No " & aOwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Migration Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Migration")
            Exit Sub
        End Try
        Try
            log.Debug("SortData_exec - " & "processing data - disabling events, screen update, cursor")
            app.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            app.EnableEvents = False
            app.ScreenUpdating = False
            i = aLOff + 1
            aLastRow = aOWs.Cells(aOWs.Cells.Rows.Count, aColDelData).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            If aLastRow <= aLOff + 1 Then
                log.Debug("SortData_exec - " & "last row less or equal first data line - exit")
                Exit Sub
            End If
            Dim jMax As Integer = aColFrom - 1
            Do
                aNameDict.Add(CStr(aOWs.Cells(aLOffTNames + 1, jMax + 1).value), jMax + 1)
                jMax += 1
            Loop While CStr(aOWs.Cells(aLOffTNames + 1, jMax + 1).value) <> "" And jMax <= aColTo
            If jMax < aColFrom Then
                log.Debug("SortData_exec - " & "jMax less then aColFrom - exit")
                Exit Sub
            End If
            Dim aRange As Excel.Range
            aRange = aOWs.Range(aOWs.Cells(aLOff + 1, aColFrom), aOWs.Cells(aLastRow, jMax))
            Dim aSortStringArray() As String
            Dim aSortField As String
            aSortStringArray = Split(aSortString, ";")
            Dim aCol As UInt64
            aOWs.Sort.SortFields.Clear()
            Dim aKey1 As Object = Nothing
            Dim aKey2 As Object = Nothing
            Dim aKey3 As Object = Nothing
            i = 1
            For Each aSortField In aSortStringArray
                If aNameDict.ContainsKey(aSortField) Then
                    aCol = aNameDict(aSortField)
                    If i = 1 Then
                        aKey1 = aOWs.Range(aOWs.Cells(aLOff + 1, aCol), aOWs.Cells(aLOff + 1, aCol))
                    ElseIf i = 2 Then
                        aKey2 = aOWs.Range(aOWs.Cells(aLOff + 1, aCol), aOWs.Cells(aLOff + 1, aCol))
                    ElseIf i = 3 Then
                        aKey3 = aOWs.Range(aOWs.Cells(aLOff + 1, aCol), aOWs.Cells(aLOff + 1, aCol))
                    End If
                    i += 1
                End If
            Next
            aRange.Sort(Key1:=aKey1, Order1:=XlSortOrder.xlAscending, Key2:=aKey2, Order2:=XlSortOrder.xlAscending, Key3:=aKey3, Order3:=XlSortOrder.xlAscending, Header:=XlYesNoGuess.xlNo, Orientation:=XlSortOrientation.xlSortColumns)
            app.StatusBar = "Sort completed"
            app.EnableEvents = True
            app.ScreenUpdating = True
            app.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            app.EnableEvents = True
            app.ScreenUpdating = True
            app.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SortData_exec failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Migration")
            log.Error("SortData_exec - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Private Sub RemoveDoulicates_exec(ByRef pIntPar As SC.TStr, Optional pNr As String = "")
        Dim aWB As Excel.Workbook
        Dim aOWs As Excel.Worksheet
        Dim aOwsName As String = If(pIntPar.value("DOUB" & pNr, "WS_DATA") <> "", pIntPar.value("SORT" & pNr, "WS_DATA"), "Data")
        Dim aLOff As Integer = If(pIntPar.value("DOUB" & pNr, "LOFF_DATA") <> "", CInt(pIntPar.value("DOUB" & pNr, "LOFF_DATA")), 4)
        Dim aLOffTNames As Integer = If(pIntPar.value("DOUB" & pNr, "LOFF_TNAMES") <> "", CInt(pIntPar.value("DOUB" & pNr, "LOFF_TNAMES")), aLOff - 1)
        Dim aColDelData As Integer = If(pIntPar.value("DOUB" & pNr, "DATA_CHEKCOL") <> "", CInt(pIntPar.value("DOUB" & pNr, "DATA_CHEKCOL")), 1)
        Dim aColFrom As Integer = If(pIntPar.value("DOUB" & pNr, "DATA_COLFROM") <> "", CInt(pIntPar.value("DOUB" & pNr, "DATA_COLFROM")), 1)
        Dim aColTo As Integer = If(pIntPar.value("DOUB" & pNr, "DATA_COLTO") <> "", CInt(pIntPar.value("DOUB" & pNr, "DATA_COLTO")), 100)
        Dim aKeyString As String = If(pIntPar.value("DOUB" & pNr, "FIELDS") <> "", pIntPar.value("DOUB" & pNr, "FIELDS"), "")
        Dim i As UInt32
        Dim aLastRow As UInt64
        Dim aNameDict As New Dictionary(Of String, UInt64)
        If aKeyString = "" Then
            log.Debug("RemoveDoulicates_exec - " & "empty field string - exit")
            Exit Sub
        End If
        aWB = app.ActiveWorkbook
        Try
            aOWs = aWB.Worksheets(aOwsName)
        Catch Exc As System.Exception
            app.EnableEvents = True
            app.ScreenUpdating = True
            app.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("No " & aOwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Migration Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Migration")
            Exit Sub
        End Try
        Try
            log.Debug("RemoveDoulicates_exec - " & "processing data - disabling events, screen update, cursor")
            app.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            app.EnableEvents = False
            app.ScreenUpdating = False
            i = aLOff + 1
            aLastRow = aOWs.Cells(aOWs.Cells.Rows.Count, aColDelData).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            If aLastRow <= aLOff + 1 Then
                log.Debug("RemoveDoulicates_exec - " & "last row less or equal first data line - exit")
                Exit Sub
            End If
            Dim jMax As Integer = aColFrom - 1
            Do
                aNameDict.Add(CStr(aOWs.Cells(aLOffTNames + 1, jMax + 1).value), jMax + 1)
                jMax += 1
            Loop While CStr(aOWs.Cells(aLOffTNames + 1, jMax + 1).value) <> "" And jMax <= aColTo
            If jMax < aColFrom Then
                log.Debug("RemoveDoulicates_exec - " & "jMax less then aColFrom - exit")
                Exit Sub
            End If
            Dim aRange As Excel.Range
            aRange = aOWs.Range(aOWs.Cells(aLOff + 1, aColFrom), aOWs.Cells(aLastRow, jMax))
            Dim aKey As String
            Dim aKeyStringArray() As String
            aKeyStringArray = Split(aKeyString, ";")
            Dim aColArray() As Object = {}
            Dim aCol As UInt64
            i = 1
            For Each aKey In aKeyStringArray
                If aNameDict.ContainsKey(aKey) Then
                    aCol = aNameDict(aKey)
                    If aColArray IsNot Nothing Then
                        Array.Resize(aColArray, aColArray.Length + 1)
                        aColArray(aColArray.Length - 1) = aCol
                    Else
                        ReDim aColArray(0)
                        aColArray(0) = aCol
                    End If
                End If
            Next
            aRange.RemoveDuplicates(Columns:=aColArray, Header:=XlYesNoGuess.xlNo)
            app.StatusBar = "Remove Doublicates completed"
            app.EnableEvents = True
            app.ScreenUpdating = True
            app.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            app.EnableEvents = True
            app.ScreenUpdating = True
            app.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("RemoveDoulicates_exec failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap Migration")
            log.Error("RemoveDoulicates_exec - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

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
