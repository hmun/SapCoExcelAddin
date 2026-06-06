' Copyright 2016-2025 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports System.Configuration
Imports System.Environment
Imports System.Uri
Imports System.IO
Imports SC = SAPCommon

Public Class MigHelperVsto
    Dim app As Microsoft.Office.Interop.Excel.Application = Globals.ThisAddIn.Application
    Public mig As SC.Migration
    Public mh As SC.MigHelper
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Sub New(ByRef pIntPar As SC.TStr, pNr As String, Optional pUselocal As Boolean = False)
        Dim aWs As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aConfigDic As New Dictionary(Of String, Object(,))
        ' setup the migration engine
        If pUselocal Then
            log.Debug("New - " & "No config file found looking for config worksheets")
            Dim aRwsName As String = If(pIntPar.value("GEN" & pNr, "WS_RULES") <> "", pIntPar.value("GEN" & pNr, "WS_RULES"), "Rules")
            Dim aPwsName As String = If(pIntPar.value("GEN" & pNr, "WS_PATTERN") <> "", pIntPar.value("GEN" & pNr, "WS_PATTERN"), "Pattern")
            Dim aCwsName As String = If(pIntPar.value("GEN" & pNr, "WS_CONSTANT") <> "", pIntPar.value("GEN" & pNr, "WS_CONSTANT"), "Constant")
            Dim aMwsName As String = If(pIntPar.value("GEN" & pNr, "WS_MAPPING") <> "", pIntPar.value("GEN" & pNr, "WS_MAPPING"), "Mapping")
            Dim aFwsName As String = If(pIntPar.value("GEN" & pNr, "WS_FORMULA") <> "", pIntPar.value("GEN" & pNr, "WS_FORMULA"), "Formula")
            ' try to read the rules from the excel workbook
            Dim i As Integer
            Dim aLastRow As Integer
            aWB = app.ActiveWorkbook
            Try
                aWs = aWB.Worksheets(aRwsName)
                i = 2
                aLastRow = 0
                Do While CStr(aWs.Cells(i, 1).value) <> ""
                    i += 1
                Loop
                aLastRow = i - 1
                Dim aRange As Excel.Range = aWs.Range(aWs.Cells(2, 1), aWs.Cells(aLastRow, 6))
                Dim aArray As Object(,) = CType(aRange.Value, Object(,))
                aConfigDic.Add(aRwsName, aArray)
            Catch Exc As System.Exception
                MsgBox("No " & aRwsName & " Rules Sheet in current workbook.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap, MigHelper")
            End Try
            Try
                aWs = aWB.Worksheets(aPwsName)
                i = 2
                aLastRow = 0
                Do While CStr(aWs.Cells(i, 1).value) <> ""
                    i += 1
                Loop
                aLastRow = i - 1
                Dim aRange As Excel.Range = aWs.Range(aWs.Cells(2, 1), aWs.Cells(aLastRow, 3))
                Dim aArray As Object(,) = CType(aRange.Value, Object(,))
                aConfigDic.Add(aPwsName, aArray)
            Catch Exc As System.Exception
                log.Debug("New - " & "No " & aPwsName & " Sheet in current workbook.")
            End Try
            Try
                aWs = aWB.Worksheets(aCwsName)
                i = 2
                aLastRow = 0
                Do While CStr(aWs.Cells(i, 1).value) <> ""
                    i += 1
                Loop
                aLastRow = i - 1
                Dim aRange As Excel.Range = aWs.Range(aWs.Cells(2, 1), aWs.Cells(aLastRow, 3))
                Dim aArray As Object(,) = CType(aRange.Value, Object(,))
                aConfigDic.Add(aCwsName, aArray)
            Catch Exc As System.Exception
                log.Debug("New - " & "No " & aCwsName & " Sheet in current workbook.")
            End Try
            Try
                aWs = aWB.Worksheets(aFwsName)
                i = 2
                aLastRow = 0
                Do While CStr(aWs.Cells(i, 1).value) <> ""
                    i += 1
                Loop
                aLastRow = i - 1
                Dim aRange As Excel.Range = aWs.Range(aWs.Cells(2, 1), aWs.Cells(aLastRow, 3))
                Dim aArray As Object(,) = CType(aRange.Value, Object(,))
                aConfigDic.Add(aFwsName, aArray)
            Catch Exc As System.Exception
                log.Debug("New - " & "No " & aFwsName & " Sheet in current workbook.")
            End Try
            Try
                aWs = aWB.Worksheets(aMwsName)
                i = 2
                aLastRow = 0
                Do While CStr(aWs.Cells(i, 1).value) <> ""
                    i += 1
                Loop
                aLastRow = i - 1
                Dim aRange As Excel.Range = aWs.Range(aWs.Cells(2, 1), aWs.Cells(aLastRow, 4))
                Dim aArray As Object(,) = CType(aRange.Value, Object(,))
                aConfigDic.Add(aMwsName, aArray)
            Catch Exc As System.Exception
                log.Debug("New - " & "No " & aMwsName & " Sheet in current workbook.")
            End Try
        End If
        mh = New SC.MigHelper(aConfigDic, pIntPar, pNr, pUselocal)
        mig = mh.mig
    End Sub

    Public Sub writeOutData(ByRef pWs As Excel.Worksheet, pTOutData As SC.TOutData, pLine As UInt64, ByRef pKeyArray As Object(,), Optional pType As String = "V", Optional pStyle As String = "")
        Dim aKvB As KeyValuePair(Of String, SortedDictionary(Of UInt64, Object))
        '        Dim aTDataList As List(Of Object)
        Dim aCol As UInt64
        Dim aColName As String
        Dim aArray(,) As Object
        Dim aSylesDic As New Dictionary(Of String, String)
        aSylesDic = mig.TargetStyles()
        For Each aKvB In pTOutData.aTDataDic
            aCol = CInt(aKvB.Key)
            Try
                aColName = pKeyArray(1, aCol)
            Catch ex As Exception
                aColName = ""
            End Try
            app.StatusBar = "Converting Column " & aCol & " to Array"
            aArray = pTOutData.toArray(aKvB.Key)
            app.StatusBar = "Writing Column " & aCol & " to Target Sheet"
            If Not aArray Is Nothing Then
                Dim aLen As UInt64 = aArray.GetLength(0)
                Dim aRange = pWs.Range(pWs.Cells(pLine, aCol), pWs.Cells(pLine + aLen - 1, aCol))
                If aSylesDic.ContainsKey(aColName) Then
                    aRange.Style = aSylesDic(aColName)
                End If
                If pType = "V" Then
                    aRange.Value = aArray
                ElseIf pType = "F" Then
                    aRange.FormulaR1C1 = aArray
                End If
            End If
        Next
    End Sub

End Class
