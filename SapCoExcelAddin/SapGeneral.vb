' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports System.Reflection
Imports System.Diagnostics

Public Class SapGeneral
    Public Function checkVersion() As Integer
        Dim aCws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aFromVersion As String
        Dim aToVersion As String
        Dim assembly As Assembly
        Dim fileVersionInfo As FileVersionInfo
        Dim aVersion As String

        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aCws = aWB.Worksheets("SAP-Con")
        Catch Exc As System.Exception
            MsgBox("No SAP-Con Sheet in current workbook. Check if the current workbook is a valid SAP Controlling Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapGeneral")
            checkVersion = False
            Exit Function
        End Try
        aFromVersion = aCws.Cells(13, 2).Value
        aToVersion = aCws.Cells(14, 2).Value

        assembly = System.Reflection.Assembly.GetExecutingAssembly()
        fileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location)
        aVersion = fileVersionInfo.ProductVersion

        If aVersion > aToVersion Or aVersion < aFromVersion Then
            MsgBox("The Version of the Excel-Template is not valid for this Add-In. Please use a Template that is valid for version " & aVersion,
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapGeneral")
            checkVersion = False
            Exit Function
        End If

        checkVersion = True
    End Function

End Class
