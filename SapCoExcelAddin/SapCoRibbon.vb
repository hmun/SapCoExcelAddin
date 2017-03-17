' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports Microsoft.Office.Tools.Ribbon
Imports SAP.Middleware.Connector

Public Class SapCoRibbon
    Private aSapCon
    Private aSapGeneral

    Private aCoAre As String
    Private aFiscy As String
    Private aPfrom As String
    Private aPto As String
    Private aSVers As String
    Private aTVers As String
    Private aCurt As String
    Private aCompCodes As String

    Private Function getParameters() As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP CO-OM Planning Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getParameters = False
            Exit Function
        End Try
        aCoAre = CStr(aPws.Cells(2, 2).Value)
        aFiscy = CStr(aPws.Cells(3, 2).Value)
        aPfrom = CStr(aPws.Cells(4, 2).Value)
        aPto = CStr(aPws.Cells(5, 2).Value)
        aSVers = CStr(aPws.Cells(6, 2).Value)
        aTVers = CStr(aPws.Cells(7, 2).Value)
        aCurt = CStr(aPws.Cells(8, 2).Value)
        aCompCodes = CStr(aPws.Cells(9, 2).Value)
        If aCoAre = "" Or
            aFiscy = "" Or
            aPfrom = "" Or
            aPto = "" Or
            aSVers = "" Or
            aTVers = "" Or
            aCurt = "" Then
            MsgBox("Please fill all obligatory fields in the parameter sheet!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getParameters = False
            Exit Function
        End If
        getParameters = True
    End Function

    Private Function checkCon() As Integer
        Dim aSapConRet As Integer
        checkCon = False
        If Not aSapGeneral.checkVersion() Then
            Exit Function
        End If
        aSapConRet = 0
        If aSapCon Is Nothing Then
            aSapCon = New SapCon
            aSapConRet = aSapCon.checkCon()
        End If
        If aSapConRet = 0 Then
            checkCon = True
        End If
    End Function

    Private Sub ButtonLogoff_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogoff.Click
        If Not aSapCon Is Nothing Then
            aSapCon.SAPlogoff()
            aSapCon = Nothing
        End If
    End Sub

    Private Sub ButtonLogon_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogon.Click
        Dim aConRet As Integer

        If Not aSapGeneral.checkVersion() Then
            Exit Sub
        End If
        If aSapCon Is Nothing Then
            aSapCon = New SapCon
        End If
        aConRet = aSapCon.checkCon()
        If aConRet = 0 Then
            MsgBox("SAP-Logon successful! ", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Sap Accounting")
        End If
    End Sub

    Private Sub ButtonReadAO_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonReadAO.Click
        Dim aSAPCOObject As New SAPCOObject
        Dim aCompCodeSplit
        Dim aCompCode
        Dim aContrl As New Collection
        Dim aData As New Collection
        Dim aObjects As New Collection
        Dim aRetStr As String
        Dim i As Integer
        Dim aRange As Excel.Range

        If getParameters() = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        aCompCodeSplit = Split(aCompCodes, ";")
        Dim aSAPGetCOObject As New SAPGetCOObject(aSapCon)
        For Each aCompCode In aCompCodeSplit
            aSAPGetCOObject.GetCoObjects("O", aFiscy, aSVers, aCoAre, CStr(aCompCode), aObjects)
        Next aCompCode
        If aObjects.Count = 0 Then
            Exit Sub
        End If
        Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
        aRetStr = aSAPCostActivityPlanning.ReadActivityOutputTot(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData, aContrl)
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("AOData")
            aDws.Activate()
            If CStr(aDws.Cells(2, 1).Value) <> "" Then
                aRange = aDws.Range("A2")
                i = 2
                Do
                    i = i + 1
                Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
                aRange = aDws.Range(aRange, aDws.Cells(i, 1))
                aRange.EntireRow.Delete()
            End If
            Dim aSapDataRow As Object
            Dim aSapContrlRow As Object
            Dim aCells As Excel.Range
            i = 1
            If aData.Count > 0 Then
                Do
                    aSapDataRow = aData(i)
                    aSapContrlRow = aContrl(i)
                    aDws.Cells(i + 1, 1) = aObjects(i).Costcenter
                    aDws.Cells(i + 1, 2) = aObjects(i).Acttype
                    aDws.Cells(i + 1, 3) = CStr(aSapDataRow.GetValue("UNIT_OF_MEASURE"))
                    aDws.Cells(i + 1, 4) = CStr(aSapDataRow.GetValue("aCURRENCY"))
                    aDws.Cells(i + 1, 5) = CDbl(aSapDataRow.GetValue("ACTVTY_QTY"))
                    aDws.Cells(i + 1, 6) = CStr(aSapDataRow.GetValue("DIST_KEY_QUAN"))
                    aDws.Cells(i + 1, 7) = CDbl(aSapDataRow.GetValue("ACTVTY_CAPACTY"))
                    aDws.Cells(i + 1, 8) = CStr(aSapDataRow.GetValue("DIST_KEY_CAPCTY"))
                    aDws.Cells(i + 1, 9) = CDbl(aSapDataRow.GetValue("PRICE_FIX"))
                    aDws.Cells(i + 1, 10) = CStr(aSapDataRow.GetValue("DIST_KEY_PRICE_FIX"))
                    aDws.Cells(i + 1, 11) = CDbl(aSapDataRow.GetValue("PRICE_VAR"))
                    aDws.Cells(i + 1, 12) = CStr(aSapDataRow.GetValue("DIST_KEY_PRICE_VAR"))
                    aDws.Cells(i + 1, 13) = CInt(aSapDataRow.GetValue("PRICE_UNIT"))
                    aDws.Cells(i + 1, 14) = CStr(aSapDataRow.GetValue("EQUIVALENCE_NO"))

                    aDws.Cells(i + 1, 15) = CStr(aSapContrlRow.GetValue("PRICE_INDICATOR"))
                    aDws.Cells(i + 1, 16) = CStr(aSapContrlRow.GetValue("SWITCH_LAYOUT"))
                    aDws.Cells(i + 1, 17) = CStr(aSapContrlRow.GetValue("ALLOC_COST_ELEM"))
                    aDws.Cells(i + 1, 18) = CInt(aSapContrlRow.GetValue("ATTRIB_INDEX"))
                    aDws.Cells(i + 1, 19) = CInt(aSapDataRow.GetValue("VALUE_INDEX"))
                    i = i + 1
                Loop While i <= aObjects.Count
            End If
            aDws.Cells(i + 1, 2) = aRetStr
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonReadAO_Click")
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonReadPC_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonReadPC.Click
        Dim aSAPCOObject As New SAPCOObject
        Dim aCompCodeSplit
        Dim aCompCode
        Dim aData As New Collection
        Dim aObjects As New Collection
        Dim aRetStr As String
        Dim i As Integer
        Dim aRange As Excel.Range

        If getParameters() = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        aCompCodeSplit = Split(aCompCodes, ";")
        Dim aSAPGetCOObject As New SAPGetCOObject(aSapCon)
        For Each aCompCode In aCompCodeSplit
            aSAPGetCOObject.GetCoObjects("P", aFiscy, aSVers, aCoAre, CStr(aCompCode), aObjects)
        Next aCompCode
        If aObjects.Count = 0 Then
            Exit Sub
        End If
        Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
        aRetStr = aSAPCostActivityPlanning.ReadPrimCostTot(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData)
        Try
            Dim aDws As Excel.Worksheet
            Dim aWB As Excel.Workbook
            aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
            aDws = aWB.Worksheets("PData")
            aDws.Activate()
            If CStr(aDws.Cells(2, 1).Value) <> "" Then
                aRange = aDws.Range("A2")
                i = 2
                Do
                    i = i + 1
                Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
                aRange = aDws.Range(aRange, aDws.Cells(i, 1))
                aRange.EntireRow.Delete()
            End If
            Dim aSapDataRow As Object
            Dim aCells As Excel.Range
            i = 1
            If aData.Count > 0 Then
                Do
                    aSapDataRow = aData(i)
                    aCells = aDws.Range(aDws.Cells(i, 1), aDws.Cells(i, 4))
                    aCells.NumberFormat = "@"
                    aDws.Cells(i + 1, 1) = aObjects(i).Costcenter
                    aDws.Cells(i + 1, 2) = aObjects(i).WBS_ELEMENT
                    aDws.Cells(i + 1, 3) = aObjects(i).Acttype
                    aDws.Cells(i + 1, 4) = aObjects(i).Costelem
                    aDws.Cells(i + 1, 5) = CStr(aSapDataRow.GetValue("TRANS_CURR"))
                    aDws.Cells(i + 1, 6) = CDbl(aSapDataRow.GetValue("FIX_VALUE"))
                    aDws.Cells(i + 1, 7) = CStr(aSapDataRow.GetValue("DIST_KEY_FIX_VAL"))
                    aDws.Cells(i + 1, 8) = CDbl(aSapDataRow.GetValue("VAR_VALUE"))
                    aDws.Cells(i + 1, 9) = CStr(aSapDataRow.GetValue("DIST_KEY_VAR_VAL"))
                    aDws.Cells(i + 1, 10) = CDbl(aSapDataRow.GetValue("FIX_QUAN"))
                    aDws.Cells(i + 1, 11) = CStr(aSapDataRow.GetValue("DIST_KEY_FIX_QUAN"))
                    aDws.Cells(i + 1, 12) = CDbl(aSapDataRow.GetValue("VAR_QUAN"))
                    aDws.Cells(i + 1, 13) = CStr(aSapDataRow.GetValue("DIST_KEY_VAR_QUAN"))
                    i = i + 1
                Loop While i <= aObjects.Count
            End If
            aDws.Cells(i + 1, 2) = aRetStr
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonReadPC_Click")
            Exit Sub
        End Try
    End Sub

    Private Sub SapCoRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        aSapGeneral = New SapGeneral
    End Sub

    Private Sub ButtonPostPC_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostPC.Click
        Dim i As Integer
        Dim aData As New Collection
        Dim aDataRow As New Collection
        Dim aObjects As New Collection
        Dim aVal
        Dim aRetStr As String

        If getParameters() = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("PData")
        Catch Exc As System.Exception
            MsgBox("No PData Sheet in current workbook. Check if the current workbook is a valid SAP CO-OM Planning Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            Exit Sub
        End Try
        aRetStr = ""
        aDws.Activate()
        Try
            i = 2
            Do
                Dim aSAPCOObject = New SAPCOObject
                aSAPCOObject = aSAPCOObject.create(CStr(aDws.Cells(i, 1).Value),
                                               CStr(aDws.Cells(i, 3).Value),
                                               CStr(aDws.Cells(i, 4).Value), "", "",
                                               CStr(aDws.Cells(i, 2).Value))
                aObjects.Add(aSAPCOObject)
                aDataRow = New Collection
                For J = 6 To 14
                    aVal = aDws.Cells(i, J).Value
                    aDataRow.Add(aVal)
                Next J
                aData.Add(aDataRow)
                i = i + 1
            Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
            Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
            aRetStr = aSAPCostActivityPlanning.PostPrimCostTot(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData)
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonPostPC_Click")
        End Try
        aDws.Cells(i, 2) = aRetStr
    End Sub

    Private Sub ButtonPostAO_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostAO.Click
        Dim i As Integer
        Dim aData As New Collection
        Dim aContrl As New Collection
        Dim aDataRow As New Collection
        Dim aContrlRow As New Collection
        Dim aObjects As New Collection
        Dim aVal
        Dim aRetStr As String

        If getParameters() = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("AOData")
        Catch Exc As System.Exception
            MsgBox("No AOData Sheet in current workbook. Check if the current workbook is a valid SAP CO-OM Planning Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            Exit Sub
        End Try
        aRetStr = ""
        aDws.Activate()
        Try
            i = 2
            Do
                Dim aSAPCOObject = New SAPCOObject
                aSAPCOObject = aSAPCOObject.create(CStr(aDws.Cells(i, 1).Value), CStr(aDws.Cells(i, 2).Value), "")
                aObjects.Add(aSAPCOObject)
                aDataRow = New Collection
                For J = 3 To 14
                    aVal = aDws.Cells(i, J).Value
                    aDataRow.Add(aVal)
                Next J
                aData.Add(aDataRow)
                For J = 15 To 17
                    aVal = aDws.Cells(i, J).Value
                    aContrlRow.Add(aVal)
                Next J
                aContrl.Add(aContrlRow)
                i = i + 1
            Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
            Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
            aRetStr = aSAPCostActivityPlanning.PostActivityOutputTot(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData, aContrl)
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonPostAO_Click")
        End Try
        aDws.Cells(i, 2) = aRetStr
    End Sub

    Private Sub ButtonPostAI_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostAI.Click
        Dim i As Integer
        Dim aData As New Collection
        Dim aDataRow As New Collection
        Dim aObjects As New Collection
        Dim aVal
        Dim aRetStr As String

        If getParameters() = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("AIData")
        Catch Exc As System.Exception
            MsgBox("No AIData Sheet in current workbook. Check if the current workbook is a valid SAP CO-OM Planning Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            Exit Sub
        End Try
        aRetStr = ""
        aDws.Activate()
        Try
            i = 2
            Do
                Dim aSAPCOObject = New SAPCOObject
                aSAPCOObject = aSAPCOObject.create(CStr(aDws.Cells(i, 1).Value),
                                               CStr(aDws.Cells(i, 3).Value), "",
                                               CStr(aDws.Cells(i, 4).Value), CStr(aDws.Cells(i, 5).Value), CStr(aDws.Cells(i, 2).Value))
                aObjects.Add(aSAPCOObject)
                aDataRow = New Collection
                For J = 6 To 10
                    aVal = aDws.Cells(i, J).Value
                    aDataRow.Add(aVal)
                Next J
                aData.Add(aDataRow)
                i = i + 1
            Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
            Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
            aRetStr = aSAPCostActivityPlanning.PostActivityInputTot(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData)
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonPostAI_Click")
        End Try
        aDws.Cells(i, 2) = aRetStr
    End Sub

    Private Sub ButtonReadAI_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonReadAI.Click
        Dim aSAPCOObject As New SAPCOObject
        Dim aCompCodeSplit
        Dim aCompCode
        Dim aData As New Collection
        Dim aContrl As New Collection
        Dim aObjects As New Collection
        Dim aRetStr As String
        Dim i As Integer
        Dim aRange As Excel.Range

        If getParameters() = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        aCompCodeSplit = Split(aCompCodes, ";")
        Dim aSAPGetCOObject As New SAPGetCOObject(aSapCon)
        For Each aCompCode In aCompCodeSplit
            aSAPGetCOObject.GetCoObjects("I", aFiscy, aSVers, aCoAre, CStr(aCompCode), aObjects)
        Next aCompCode
        If aObjects.Count = 0 Then
            Exit Sub
        End If
        Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
        aRetStr = aSAPCostActivityPlanning.ReadActivityOutputTot(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData, aContrl)
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("AIData")
            aDws.Activate()
            If CStr(aDws.Cells(2, 1).Value) <> "" Then
                aRange = aDws.Range("A2")
                i = 2
                Do
                    i = i + 1
                Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
                aRange = aDws.Range(aRange, aDws.Cells(i, 1))
                aRange.EntireRow.Delete()
            End If
            Dim aSapDataRow As Object
            i = 1
            If aData.Count > 0 Then
                Do
                    aSapDataRow = aData(i)
                    aDws.Cells(i + 1, 1) = aObjects(i).Costcenter
                    aDws.Cells(i + 1, 2) = aObjects(i).WBS_ELEMENT
                    aDws.Cells(i + 1, 3) = aObjects(i).Acttype
                    aDws.Cells(i + 1, 4) = aObjects(i).SCostcenter
                    aDws.Cells(i + 1, 5) = aObjects(i).SActtype
                    aDws.Cells(i + 1, 6) = CStr(aSapDataRow.GetValue("UNIT_OF_MEASURE"))
                    aDws.Cells(i + 1, 7) = CDbl(aSapDataRow.GetValue("QUANTITY_FIX"))
                    aDws.Cells(i + 1, 8) = CStr(aSapDataRow.GetValue("DIST_KEY_QUAN_FIX"))
                    aDws.Cells(i + 1, 9) = CDbl(aSapDataRow.GetValue("QUANTITY_VAR"))
                    aDws.Cells(i + 1, 10) = CStr(aSapDataRow.GetValue("DIST_KEY_QUAN_VAR"))
                    i = i + 1
                Loop While i <= aObjects.Count
            End If
            aDws.Cells(i + 1, 2) = aRetStr
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonReadAI_Click")
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonReadSK_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonReadSK.Click
        Dim aSAPCOObject As New SAPCOObject
        Dim aCompCodeSplit
        Dim aCompCode
        Dim aData As New Collection
        Dim aObjects As New Collection
        Dim aRetStr As String
        Dim i As Integer
        Dim aRange As Excel.Range

        If getParameters() = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        aCompCodeSplit = Split(aCompCodes, ";")
        Dim aSAPGetCOObject As New SAPGetCOObject(aSapCon)
        For Each aCompCode In aCompCodeSplit
            '   TODO change that to read the objects with key-figure plan
            aSAPGetCOObject.GetCoObjects("O", aFiscy, aSVers, aCoAre, CStr(aCompCode), aObjects)
        Next aCompCode
        If aObjects.Count = 0 Then
            Exit Sub
        End If
        Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
        aRetStr = aSAPCostActivityPlanning.ReadKeyFigure(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData)
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("SKData")
            aDws.Activate()
            If CStr(aDws.Cells(2, 1).Value) <> "" Then
                aRange = aDws.Range("A2")
                i = 2
                Do
                    i = i + 1
                Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
                aRange = aDws.Range(aRange, aDws.Cells(i, 1))
                aRange.EntireRow.Delete()
            End If
            Dim aSapDataRow As Object
            i = 1
            If aData.Count > 0 Then
                Do
                    aSapDataRow = aData(i)
                    aDws.Cells(i + 1, 1) = aObjects(i).Costcenter
                    aDws.Cells(i + 1, 2) = aObjects(i).WBS_ELEMENT
                    aDws.Cells(i + 1, 3) = aObjects(i).Acttype
                    aDws.Cells(i + 1, 4) = CStr(aSapDataRow("STATKEYFIG"))
                    aDws.Cells(i + 1, 5) = CStr(aSapDataRow("UNIT_OF_MEASURE"))
                    aDws.Cells(i + 1, 6) = CDbl(aSapDataRow("QUANTITY_PER01"))
                    aDws.Cells(i + 1, 7) = CDbl(aSapDataRow("QUANTITY_PER02"))
                    aDws.Cells(i + 1, 8) = CDbl(aSapDataRow("QUANTITY_PER03"))
                    aDws.Cells(i + 1, 9) = CDbl(aSapDataRow("QUANTITY_PER04"))
                    aDws.Cells(i + 1, 10) = CDbl(aSapDataRow("QUANTITY_PER05"))
                    aDws.Cells(i + 1, 11) = CDbl(aSapDataRow("QUANTITY_PER06"))
                    aDws.Cells(i + 1, 12) = CDbl(aSapDataRow("QUANTITY_PER07"))
                    aDws.Cells(i + 1, 13) = CDbl(aSapDataRow("QUANTITY_PER08"))
                    aDws.Cells(i + 1, 14) = CDbl(aSapDataRow("QUANTITY_PER09"))
                    aDws.Cells(i + 1, 15) = CDbl(aSapDataRow("QUANTITY_PER010"))
                    aDws.Cells(i + 1, 16) = CDbl(aSapDataRow("QUANTITY_PER011"))
                    aDws.Cells(i + 1, 17) = CDbl(aSapDataRow("QUANTITY_PER012"))
                    i = i + 1
                Loop While i <= aObjects.Count
            End If
            aDws.Cells(i + 1, 2) = aRetStr
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonReadSK_Click")
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonPostSK_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostSK.Click
        Dim i As Integer
        Dim aData As New Collection
        Dim aDataRow As New Collection
        Dim aObjects As New Collection
        Dim aVal
        Dim aRetStr As String

        If getParameters() = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("SKData")
        Catch Exc As System.Exception
            MsgBox("No SKData Sheet in current workbook. Check if the current workbook is a valid SAP CO-OM Planning Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            Exit Sub
        End Try
        aRetStr = ""
        aDws.Activate()
        Try
            i = 2
            Do
                Dim aSAPCOObject = New SAPCOObject
                aSAPCOObject = aSAPCOObject.create(CStr(aDws.Cells(i, 1).Value), CStr(aDws.Cells(i, 3).Value), "", "", "", CStr(aDws.Cells(i, 2).Value))
                aObjects.Add(aSAPCOObject)
                aDataRow = New Collection
                For J = 6 To 17
                    aVal = aDws.Cells(i, J).Value
                    aDataRow.Add(aVal)
                Next J
                aData.Add(aDataRow)
                i = i + 1
            Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
            Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
            aRetStr = aSAPCostActivityPlanning.PostKeyFigure(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData)
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonPostSK_Click")
        End Try
        aDws.Cells(i, 2) = aRetStr
    End Sub
End Class
