' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports System.Configuration
Imports System.Environment
Imports System.Uri
Imports System.IO
Imports SAPCommon

Public Class MigHelper
    Public mig As SAPCommon.Migration
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private aFilterField As String = ""
    Private aFilterOperation As String = ""
    Private aFilterCompare As String = ""

    Public Sub New(ByRef pPar As SAPCommon.TStr, pNr As String, Optional pBaseFilterStr As String = "")
        Dim aWs As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim configFile As String = ""
        Dim aGenLocalRules As String = If(pPar.value("GEN", "LOCAL_RULES") <> "", CStr(pPar.value("GEN", "LOCAL_RULES")), "")
        log.Debug("MigHelper.New - aGenLocalRules = " & aGenLocalRules)
        Dim aUselocal As Boolean = If(aGenLocalRules = "X", True, False)
        ' set Filter Fields
        If Not String.IsNullOrEmpty(pBaseFilterStr) Then
            Dim aFilterStr() As String = {}
            aFilterStr = pBaseFilterStr.Split(";")
            If aFilterStr.Length = 3 Then
                aFilterField = aFilterStr(0)
                aFilterOperation = aFilterStr(1)
                aFilterCompare = aFilterStr(2)
                If aFilterCompare.ToUpper() = "NULL" Then
                    aFilterCompare = ""
                End If
            End If
        End If
        ' Check for local rules first
        If Not aUselocal Then
            log.Debug("MigHelper.New - using XML rules")
            Dim assemblyName As System.Reflection.AssemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName()
            Dim assembly As String = assemblyName.Name
            Dim appData As String = GetFolderPath(Environment.SpecialFolder.ApplicationData)
            configFile = Uri.UnescapeDataString(appData & "\SapExcel\" & assembly & "\mig_rules" & pNr & ".config")
            log.Debug("New - " & "looking for config file=" & configFile)
            If Not System.IO.File.Exists(configFile) Then
                appData = GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                configFile = Uri.UnescapeDataString(appData & "\SapExcel\" & assembly & "\mig_rules" & pNr & ".config")
                log.Debug("New - " & "looking for config file=" & configFile)
                If Not System.IO.File.Exists(configFile) Then
                    appData = New Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).AbsolutePath
                    appData = Path.GetDirectoryName(appData)
                    configFile = Uri.UnescapeDataString(appData & "\mig_rules" & pNr & ".config")
                    log.Debug("New - " & "looking for config file=" & configFile)
                    If Not System.IO.File.Exists(configFile) Then
                        configFile = ""
                    End If
                End If
            End If
        End If
        ' setup the migration engine
        If Not configFile = "" Then
            log.Debug("New - " & "found config file=" & configFile)
            mig = New SAPCommon.Migration(configFile)
        Else
            log.Debug("New - " & "No config file found looking for config worksheets")
            Dim aRwsName As String = If(pPar.value("GEN" & pNr, "WS_RULES") <> "", pPar.value("GEN" & pNr, "WS_RULES"), "Rules")
            Dim aPwsName As String = If(pPar.value("GEN" & pNr, "WS_PATTERN") <> "", pPar.value("GEN" & pNr, "WS_PATTERN"), "Pattern")
            Dim aCwsName As String = If(pPar.value("GEN" & pNr, "WS_CONSTANT") <> "", pPar.value("GEN" & pNr, "WS_CONSTANT"), "Constant")
            Dim aMwsName As String = If(pPar.value("GEN" & pNr, "WS_MAPPING") <> "", pPar.value("GEN" & pNr, "WS_MAPPING"), "Mapping")
            Dim aFwsName As String = If(pPar.value("GEN" & pNr, "WS_FORMULA") <> "", pPar.value("GEN" & pNr, "WS_FORMULA"), "Formula")
            mig = New SAPCommon.Migration()
            ' try to read the rules from the excel workbook
            Dim i As Integer
            aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
            Try
                aWs = aWB.Worksheets(aRwsName)
                i = 2
                Do While CStr(aWs.Cells(i, 1).value) <> ""
                    mig.AddRule(CStr(aWs.Cells(i, 1).Value), CStr(aWs.Cells(i, 2).Value), CStr(aWs.Cells(i, 3).Value), CStr(aWs.Cells(i, 4).Value))
                    If CStr(aWs.Cells(i, 3).Value = "C" And Not String.IsNullOrEmpty(CStr(aWs.Cells(i, 5).Value))) Then
                        mig.AddConstant(CStr(aWs.Cells(i, 1).Value), CStr(aWs.Cells(i, 2).Value), CStr(aWs.Cells(i, 5).Value))
                    End If
                    If CStr(aWs.Cells(i, 3).Value = "P" And Not String.IsNullOrEmpty(CStr(aWs.Cells(i, 5).Value))) Then
                        mig.AddPattern(CStr(aWs.Cells(i, 1).Value), CStr(aWs.Cells(i, 2).Value), CStr(aWs.Cells(i, 5).Value))
                    End If
                    If CStr(aWs.Cells(i, 3).Value = "F" And Not String.IsNullOrEmpty(CStr(aWs.Cells(i, 5).Value))) Then
                        mig.AddFormula(CStr(aWs.Cells(i, 1).Value), CStr(aWs.Cells(i, 2).Value), CStr(aWs.Cells(i, 5).Value))
                    End If
                    i += 1
                Loop
            Catch Exc As System.Exception
                MsgBox("No " & aRwsName & " Rules Sheet in current workbook.",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO, MigHelper")
            End Try
            Try
                aWs = aWB.Worksheets(aPwsName)
                i = 2
                Do While CStr(aWs.Cells(i, 1).value) <> ""
                    mig.AddPattern(CStr(aWs.Cells(i, 1).Value), CStr(aWs.Cells(i, 2).Value), CStr(aWs.Cells(i, 3).Value))
                    i += 1
                Loop
            Catch Exc As System.Exception
                log.Debug("New - " & "No " & aPwsName & " Sheet in current workbook.")
            End Try
            Try
                aWs = aWB.Worksheets(aCwsName)
                i = 2
                Do While CStr(aWs.Cells(i, 1).value) <> ""
                    mig.AddConstant(CStr(aWs.Cells(i, 1).Value), CStr(aWs.Cells(i, 2).Value), CStr(aWs.Cells(i, 3).Value))
                    i += 1
                Loop
            Catch Exc As System.Exception
                log.Debug("New - " & "No " & aCwsName & " Sheet in current workbook.")
            End Try
            Try
                aWs = aWB.Worksheets(aFwsName)
                i = 2
                Do While CStr(aWs.Cells(i, 1).value) <> ""
                    mig.AddFormula(CStr(aWs.Cells(i, 1).Value), CStr(aWs.Cells(i, 2).Value), CStr(aWs.Cells(i, 3).Value))
                    i += 1
                Loop
            Catch Exc As System.Exception
                log.Debug("New - " & "No " & aFwsName & " Sheet in current workbook.")
            End Try
            Try
                aWs = aWB.Worksheets(aMwsName)
                i = 2
                Do While CStr(aWs.Cells(i, 1).value) <> ""
                    mig.AddMapping(CStr(aWs.Cells(i, 2).Value), CStr(aWs.Cells(i, 1).Value), CStr(aWs.Cells(i, 3).Value), CStr(aWs.Cells(i, 4).Value))
                    i += 1
                Loop
            Catch Exc As System.Exception
                log.Debug("New - " & "No " & aMwsName & " Sheet in current workbook.")
            End Try
        End If
    End Sub

    Sub saveToConfig(pNr As String)
        Dim assemblyName As System.Reflection.AssemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName()
        Dim assembly As String = assemblyName.Name
        Dim appData As String = GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        Dim configFile As String = appData & "\SapExcel\" & assembly & "\mig_rules" & pNr & ".config"
        Dim config As Configuration
        Dim configMap As New ExeConfigurationFileMap
        configMap.ExeConfigFilename = configFile
        config = TryCast(ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None), Configuration)
        config.Sections.Add("MigRules", mig.MRS)
        mig.MRS.SectionInformation.ForceSave = True
        config.Save(ConfigurationSaveMode.Full)
    End Sub

    Public Function isFiltered(ByRef pBaseRecord As Dictionary(Of String, SAPCommon.TField)) As Boolean
        Dim aField As SAPCommon.TField
        isFiltered = False
        If pBaseRecord.ContainsKey(aFilterField) Then
            aField = pBaseRecord(aFilterField)
            If aFilterOperation = "EQ" And aField.Value = aFilterCompare Then
                isFiltered = True
            ElseIf aFilterOperation = "NE" And aField.Value <> aFilterCompare Then
                isFiltered = True
            End If
        Else
            If aFilterOperation = "NE" And (String.IsNullOrEmpty(aFilterCompare) Or aFilterCompare = "#") Then
                isFiltered = False
            End If
        End If
    End Function

    Public Function makeDictForRules(ByRef pWs As Excel.Worksheet, pRow As Integer, pHeaderRow As Integer, pFromCol As Integer, pToCol As Integer) As Dictionary(Of String, SAPCommon.TField)
        Dim retDict As New Dictionary(Of String, SAPCommon.TField)
        Dim tfield As New SAPCommon.TField
        For j = pFromCol To pToCol
            If Not CStr(pWs.Cells(pHeaderRow, j).Value) = "" Then
                If mig.ContainsSource("P", CStr(pWs.Cells(pHeaderRow, j).Value)) Or
                   mig.ContainsSource("C", CStr(pWs.Cells(pHeaderRow, j).Value)) Then
                    tfield = New SAPCommon.TField(CStr(pWs.Cells(pHeaderRow, j).Value), CStr(pWs.Cells(pRow, j).Value))
                    retDict.Add(tfield.Name, tfield)
                End If
            End If
        Next
        makeDictForRules = retDict
    End Function
    Function makeDict(ByRef pWs As Excel.Worksheet, pRow As Integer, pHeaderRow As Integer, pFromCol As Integer, pToCol As Integer) As Dictionary(Of String, SAPCommon.TField)
        Dim retDict As New Dictionary(Of String, SAPCommon.TField)
        Dim tfield As New SAPCommon.TField
        For j = pFromCol To pToCol
            If Not CStr(pWs.Cells(pHeaderRow, j).Value) = "" Then
                tfield = New SAPCommon.TField(CStr(pWs.Cells(pHeaderRow, j).Value), CStr(pWs.Cells(pRow, j).Value))
                retDict.Add(tfield.Name, tfield)
            End If
        Next j
        makeDict = retDict
    End Function

End Class
