' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports Microsoft.Office.Tools.Ribbon
Imports SAP.Middleware.Connector

Public Class SapCoRibbon
    Private aSapCon
    Private aSapGeneral
    Const AI_CM = 19 ' column of message for activity posting
    Const PC_CM = 21 ' column of message for primary cost reposting

    Private aCoAre As String
    Private aFiscy As String
    Private aPfrom As String
    Private aPto As String
    Private aSVers As String
    Private aTVers As String
    Private aCurt As String
    Private aCompCodes As String

    Private aMaxLines As String

    Private Function getParameters(pType As String) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim akey As String
        Dim aName As String

        aName = "SAPCoOmPlanning" & pType
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP CO-OM Planning Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getParameters = False
            Exit Function
        End Try
        akey = CStr(aPws.Cells(1, 1).Value)
        If akey <> aName Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key " & aName & ". Check if the current workbook is a valid SAP CO-OM Planning Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getParameters = False
            Exit Function
        End If
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

        If getParameters("Total") = False Then
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
            i = 1
            If aData.Count > 0 Then
                Do
                    aSapDataRow = aData(i)
                    aSapContrlRow = aContrl(i)
                    aDws.Cells(i + 1, 1) = aObjects(i).Costcenter
                    aDws.Cells(i + 1, 2) = aObjects(i).Acttype
                    aDws.Cells(i + 1, 3) = CStr(aSapDataRow.GetValue("UNIT_OF_MEASURE"))
                    aDws.Cells(i + 1, 4) = CStr(aSapDataRow.GetValue("CURRENCY"))
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

        If getParameters("Total") = False Then
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

        If getParameters("Total") = False Then
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

        If getParameters("Total") = False Then
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

        If getParameters("Total") = False Then
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

        If getParameters("Total") = False Then
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
        aRetStr = aSAPCostActivityPlanning.ReadActivityInputTot(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData)
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
        ButtonReadSK_Execute("Total")
    End Sub

    Private Sub ButtonReadSK_Execute(pType As String)
        Dim aSAPCOObject As New SAPCOObject
        Dim aCompCodeSplit
        Dim aCompCode
        Dim aData As New Collection
        Dim aObjects As New Collection
        Dim aRetStr As String
        Dim i As Integer
        Dim aRange As Excel.Range

        If getParameters(pType) = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        aCompCodeSplit = Split(aCompCodes, ";")
        Dim aSAPGetCOObject As New SAPGetCOObject(aSapCon)
        For Each aCompCode In aCompCodeSplit
            aSAPGetCOObject.GetCoObjects("S", aFiscy, aSVers, aCoAre, CStr(aCompCode), aObjects)
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
                    aDws.Cells(i + 1, 4) = CStr(aSapDataRow.GetValue("STATKEYFIG"))
                    aDws.Cells(i + 1, 5) = CStr(aSapDataRow.GetValue("UNIT_OF_MEASURE"))
                    aDws.Cells(i + 1, 6) = CDbl(aSapDataRow.GetValue("QUANTITY_PER01"))
                    aDws.Cells(i + 1, 7) = CDbl(aSapDataRow.GetValue("QUANTITY_PER02"))
                    aDws.Cells(i + 1, 8) = CDbl(aSapDataRow.GetValue("QUANTITY_PER03"))
                    aDws.Cells(i + 1, 9) = CDbl(aSapDataRow.GetValue("QUANTITY_PER04"))
                    aDws.Cells(i + 1, 10) = CDbl(aSapDataRow.GetValue("QUANTITY_PER05"))
                    aDws.Cells(i + 1, 11) = CDbl(aSapDataRow.GetValue("QUANTITY_PER06"))
                    aDws.Cells(i + 1, 12) = CDbl(aSapDataRow.GetValue("QUANTITY_PER07"))
                    aDws.Cells(i + 1, 13) = CDbl(aSapDataRow.GetValue("QUANTITY_PER08"))
                    aDws.Cells(i + 1, 14) = CDbl(aSapDataRow.GetValue("QUANTITY_PER09"))
                    aDws.Cells(i + 1, 15) = CDbl(aSapDataRow.GetValue("QUANTITY_PER10"))
                    aDws.Cells(i + 1, 16) = CDbl(aSapDataRow.GetValue("QUANTITY_PER11"))
                    aDws.Cells(i + 1, 17) = CDbl(aSapDataRow.GetValue("QUANTITY_PER12"))
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
        ButtonPostSK_Execute("Total")
    End Sub

    Private Sub ButtonPostSK_Execute(pType As String)
        Dim i As Integer
        Dim aData As New Collection
        Dim aDataRow As New Collection
        Dim aObjects As New Collection
        Dim aVal
        Dim aRetStr As String

        If getParameters(pType) = False Then
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
                aSAPCOObject = aSAPCOObject.create(CStr(aDws.Cells(i, 1).Value), CStr(aDws.Cells(i, 3).Value), "", "", "",
                                                   CStr(aDws.Cells(i, 2).Value), CStr(aDws.Cells(i, 4).Value))
                aObjects.Add(aSAPCOObject)
                aDataRow = New Collection
                For J = 5 To 17
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

    Private Function getActivityAllocParameters() As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP CO ActivityAlloc Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getActivityAllocParameters = False
            Exit Function
        End Try
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> "SAPAcctngActivityAlloc" Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key SAPAcctngActivityAlloc. Check if the current workbook is a valid SAP CO Activity Allocation Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getActivityAllocParameters = False
            Exit Function
        End If
        aCoAre = CStr(aPws.Cells(2, 2).Value)
        aMaxLines = CInt(aPws.Cells(3, 2).Value)
        If aCoAre = "" Then
            MsgBox("Please fill all obligatory fields in the parameter sheet!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ActivityAlloc")
            getActivityAllocParameters = False
            Exit Function
        End If
        getActivityAllocParameters = True
    End Function

    Private Sub ButtonActivityAllocCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonActivityAllocCheck.Click
        SAP_ActivityAlloc_execute(pTest:=True)
    End Sub

    Private Sub ButtonActivityAllocPost_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonActivityAllocPost.Click
        SAP_ActivityAlloc_execute(pTest:=False)
    End Sub

    Private Sub SAP_ActivityAlloc_execute(pTest As Boolean)
        Dim i As Integer
        Dim aLines As Integer
        Dim aPostLine As Integer
        Dim aData As New Collection
        Dim aRetStr As String
        Dim aDateFormatString As New DateFormatString
        Dim aSAPAcctngActivityItem As New SAPAcctngActivityItem

        If getActivityAllocParameters() = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        Dim aSAPAcctngActivityAlloc As New SAPAcctngActivityAlloc(aSapCon)
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("Data")
        Catch Exc As System.Exception
            MsgBox("No Data Sheet in current workbook. Check if the current workbook is a valid SAP CO ActivityAlloc Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            Exit Sub
        End Try
        aRetStr = ""
        aDws.Activate()
        Dim aBUDAT As String
        Dim aBLDAT As String
        Dim aCells As Excel.Range
        aBUDAT = ""
        aBLDAT = ""
        Try
            i = 2
            aLines = 1
            aPostLine = i - 1
            Do
                If InStr(CStr(aDws.Cells(i, AI_CM).Value), "Beleg wird unter der Nummer") = 0 And
                   InStr(CStr(aDws.Cells(i, AI_CM).Value), "Document is posted under number") = 0 Then
                    If aBUDAT = "" Or aMaxLines = 1 Then
                        aBUDAT = Format(aDws.Cells(i, 1).Value, aDateFormatString.getString)
                        aBLDAT = Format(aDws.Cells(i, 2).Value, aDateFormatString.getString)
                    End If
                    aSAPAcctngActivityItem = aSAPAcctngActivityItem.create(CStr(aDws.Cells(i, 3).Value), CStr(aDws.Cells(i, 4).Value),
                                                                           CStr(aDws.Cells(i, 5).Value), CDbl(aDws.Cells(i, 6).Value),
                                                                           CStr(aDws.Cells(i, 7).Value), CStr(aDws.Cells(i, 8).Value),
                                                                           CStr(aDws.Cells(i, 9).Value), CStr(aDws.Cells(i, 10).Value),
                                                                           CStr(aDws.Cells(i, 11).Value), CStr(aDws.Cells(i, 12).Value),
                                                                           CDbl(aDws.Cells(i, 14).Value), CDbl(aDws.Cells(i, 15).Value),
                                                                           CDbl(aDws.Cells(i, 16).Value), CInt(aDws.Cells(i, 17).Value),
                                                                           CStr(aDws.Cells(i, 18).Value), CStr(aDws.Cells(i, 13).Value))
                    aData.Add(aSAPAcctngActivityItem)
                    If aLines >= CInt(aMaxLines) Then
                        aRetStr = aSAPAcctngActivityAlloc.post(aCoAre, CDate(aBUDAT), CDate(aBLDAT), aData, pTest)
                        aCells = aDws.Range(aDws.Cells(aPostLine + 1, AI_CM), aDws.Cells(i, AI_CM))
                        aCells.Value = aRetStr
                        aData = New Collection
                        aLines = 1
                        aBUDAT = ""
                        aPostLine = i
                    Else
                        aLines = aLines + 1
                    End If
                End If
                i = i + 1
            Loop While CStr(aDws.Cells(i, 1).value) <> ""
            If aData.Count > 0 Then
                aRetStr = aSAPAcctngActivityAlloc.post(aCoAre, CDate(aBUDAT), CDate(aBLDAT), aData, pTest)
                aCells = aDws.Range(aDws.Cells(aPostLine + 1, AI_CM), aDws.Cells(i - 1, AI_CM))
                aCells.Value = aRetStr
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_ActivityAlloc_execute")
        End Try

    End Sub

    Private Function getCostPostingParameters() As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP CO RepstPrimCosts Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getCostPostingParameters = False
            Exit Function
        End Try
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> "SAPAcctngRepstPrimCosts" Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key SAPAcctngActivityAlloc. Check if the current workbook is a valid SAP CO RepstPrimCosts Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getCostPostingParameters = False
            Exit Function
        End If
        aCoAre = CStr(aPws.Cells(2, 2).Value)
        aMaxLines = CInt(aPws.Cells(3, 2).Value)
        If aCoAre = "" Then
            MsgBox("Please fill all obligatory fields in the parameter sheet!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap RepstPrimCosts")
            getCostPostingParameters = False
            Exit Function
        End If
        getCostPostingParameters = True
    End Function

    Private Sub ButtonRepstPrimCostsCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonRepstPrimCostsCheck.Click
        SAP_RepstPrimCosts_execute(pTest:=True)
    End Sub

    Private Sub ButtonRepstPrimCostsPost_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonRepstPrimCostsPost.Click
        SAP_RepstPrimCosts_execute(pTest:=False)
    End Sub

    Private Sub SAP_RepstPrimCosts_execute(pTest As Boolean)
        Dim i As Integer
        Dim aLines As Integer
        Dim aPostLine As Integer
        Dim aData As New Collection
        Dim aRetStr As String
        Dim aDateFormatString As New DateFormatString
        Dim aSAPAcctngPrimCostsItem As New SAPAcctngPrimCostsItem

        If getCostPostingParameters() = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        Dim aSAPAcctngRepstPrimCosts As New SAPAcctngRepstPrimCosts(aSapCon)
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("Data")
        Catch Exc As System.Exception
            MsgBox("No Data Sheet in current workbook. Check if the current workbook is a valid SAP CO RepstPrimCosts Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            Exit Sub
        End Try
        aRetStr = ""
        aDws.Activate()
        Dim aBUDAT As String
        Dim aBLDAT As String
        Dim aCells As Excel.Range
        aBUDAT = ""
        aBLDAT = ""
        Try
            i = 2
            aLines = 1
            aPostLine = i - 1
            Do
                If InStr(CStr(aDws.Cells(i, PC_CM).Value), "Beleg wird unter der Nummer") = 0 And
                   InStr(CStr(aDws.Cells(i, PC_CM).Value), "Document is posted under number") = 0 Then
                    If aBUDAT = "" Or aMaxLines = 1 Then
                        aBUDAT = Format(aDws.Cells(i, 1).Value, aDateFormatString.getString)
                        aBLDAT = Format(aDws.Cells(i, 2).Value, aDateFormatString.getString)
                    End If
                    aSAPAcctngPrimCostsItem = aSAPAcctngPrimCostsItem.create(CStr(aDws.Cells(i, 3).Value), CStr(aDws.Cells(i, 4).Value),
                                                                             CStr(aDws.Cells(i, 5).Value), CStr(aDws.Cells(i, 6).Value),
                                                                             CStr(aDws.Cells(i, 7).Value), CStr(aDws.Cells(i, 8).Value),
                                                                             CStr(aDws.Cells(i, 9).Value), CStr(aDws.Cells(i, 10).Value),
                                                                             CStr(aDws.Cells(i, 11).Value), CDbl(aDws.Cells(i, 12).Value),
                                                                             CStr(aDws.Cells(i, 13).Value), CStr(aDws.Cells(i, 14).Value),
                                                                             CStr(aDws.Cells(i, 15).Value), CStr(aDws.Cells(i, 16).Value),
                                                                             CStr(aDws.Cells(i, 17).Value), CStr(aDws.Cells(i, 18).Value),
                                                                             CStr(aDws.Cells(i, 19).Value), CStr(aDws.Cells(i, 20).Value))
                    aData.Add(aSAPAcctngPrimCostsItem)
                    If aLines >= CInt(aMaxLines) Then
                        aRetStr = aSAPAcctngRepstPrimCosts.post(aCoAre, CDate(aBUDAT), CDate(aBLDAT), aData, pTest)
                        aCells = aDws.Range(aDws.Cells(aPostLine + 1, PC_CM), aDws.Cells(i, PC_CM))
                        aCells.Value = aRetStr
                        aData = New Collection
                        aLines = 1
                        aBUDAT = ""
                        aPostLine = i
                    Else
                        aLines = aLines + 1
                    End If
                End If
                i = i + 1
            Loop While CStr(aDws.Cells(i, 1).value) <> ""
            If aData.Count > 0 Then
                aRetStr = aSAPAcctngRepstPrimCosts.post(aCoAre, CDate(aBUDAT), CDate(aBLDAT), aData, pTest)
                aCells = aDws.Range(aDws.Cells(aPostLine + 1, PC_CM), aDws.Cells(i - 1, PC_CM))
                aCells.Value = aRetStr
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_RepstPrimCosts_execute")
        End Try
    End Sub

    Private Sub ButtonReadPerAO_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonReadPerAO.Click
        Dim aSAPCOObject As New SAPCOObject
        Dim aCompCodeSplit
        Dim aCompCode
        Dim aContrl As New Collection
        Dim aData As New Collection
        Dim aObjects As New Collection
        Dim aRetStr As String
        Dim i As Integer
        Dim aRange As Excel.Range
        Dim J As Integer
        Dim aVal

        If getParameters("Periodic") = False Then
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
        aRetStr = aSAPCostActivityPlanning.ReadActivityOutput(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData)
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
            i = 1
            If aData.Count > 0 Then
                Do
                    aSapDataRow = aData(i)
                    aDws.Cells(i + 1, 1).Value = aObjects(i).Costcenter
                    aDws.Cells(i + 1, 2).Value = aObjects(i).Acttype
                    aDws.Cells(i + 1, 3).Value = CStr(aSapDataRow.GetValue("UNIT_OF_MEASURE"))
                    aDws.Cells(i + 1, 4).Value = CStr(aSapDataRow.GetValue("CURRENCY"))
                    For J = 2 To 65
                        aVal = CDbl(aSapDataRow.GetValue(J - 1))
                        aDws.Cells(i + 1, J + 3).Value = aVal
                    Next J
                    For J = 66 To 97
                        aVal = CInt(aSapDataRow.GetValue(J - 1))
                        aDws.Cells(i + 1, J + 3).Value = aVal
                    Next J
                    i = i + 1
                Loop While i <= aObjects.Count
            End If
            aDws.Cells(i + 1, 2) = aRetStr
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonReadPerAO_Click")
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonPostPerAO_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostPerAO.Click
        Dim i As Integer
        Dim aData As New Collection
        Dim aContrl As New Collection
        Dim aDataRow As New Collection
        Dim aContrlRow As New Collection
        Dim aObjects As New Collection
        Dim aVal
        Dim aRetStr As String

        If getParameters("Periodic") = False Then
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
                For J = 2 To 65
                    aVal = aDws.Cells(i, J + 3).Value
                    aDataRow.Add(CDbl(aVal))
                Next J
                For J = 66 To 97
                    aVal = aDws.Cells(i, J + 3).Value
                    aDataRow.Add(CInt(aVal))
                Next J
                aDataRow.Add(CStr(aDws.Cells(i, 3).Value)) 'Unit of Measure
                aDataRow.Add(CStr(aDws.Cells(i, 4).Value)) 'Curr.
                aData.Add(aDataRow)
                i = i + 1
            Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
            Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
            aRetStr = aSAPCostActivityPlanning.PostActivityOutput(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData)
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonPostPerAO_Click")
        End Try
        aDws.Cells(i, 2) = aRetStr
    End Sub

    Private Sub ButtonReadPerPC_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonReadPerPC.Click
        Dim aSAPCOObject As New SAPCOObject
        Dim aCompCodeSplit
        Dim aCompCode
        Dim aData As New Collection
        Dim aObjects As New Collection
        Dim aRetStr As String
        Dim i As Integer
        Dim aRange As Excel.Range
        Dim J As Integer
        Dim aVal

        If getParameters("Periodic") = False Then
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
        aRetStr = aSAPCostActivityPlanning.ReadPrimCost(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData)
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
                    '                    aCells.NumberFormat = "@"
                    aDws.Cells(i + 1, 1).Value = aObjects(i).Costcenter
                    aDws.Cells(i + 1, 2).Value = aObjects(i).WBS_ELEMENT
                    aDws.Cells(i + 1, 3).Value = aObjects(i).Acttype
                    aDws.Cells(i + 1, 4).Value = aObjects(i).Costelem
                    aDws.Cells(i + 1, 5).Value = CStr(aSapDataRow.GetValue("TRANS_CURR"))
                    For J = 8 To 71
                        aVal = CDbl(aSapDataRow.GetValue(J - 1))
                        aDws.Cells(i + 1, J - 2).Value = aVal
                    Next J
                    i = i + 1
                Loop While i <= aObjects.Count
            End If
            aDws.Cells(i + 1, 2) = aRetStr
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonReadPerPC_Click")
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonPostPerPC_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostPerPC.Click
        Dim i As Integer
        Dim aData As New Collection
        Dim aDataRow As New Collection
        Dim aObjects As New Collection
        Dim aVal
        Dim aRetStr As String

        If getParameters("Periodic") = False Then
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
                For J = 8 To 71
                    aVal = aDws.Cells(i, J - 2).Value
                    aDataRow.Add(aVal)
                Next J
                aData.Add(aDataRow)
                i = i + 1
            Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
            Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
            aRetStr = aSAPCostActivityPlanning.PostPrimCost(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData)
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonPostPerPC_Click")
        End Try
        aDws.Cells(i, 2) = aRetStr
    End Sub

    Private Sub ButtonPostPerAI_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostPerAI.Click
        Dim i As Integer
        Dim aData As New Collection
        Dim aDataRow As New Collection
        Dim aObjects As New Collection
        Dim aVal
        Dim aRetStr As String

        If getParameters("Periodic") = False Then
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
                aDataRow.Add(CStr(aDws.Cells(i, 6).Value)) ' Unit of Meassure
                For J = 6 To 37
                    aVal = CDbl(aDws.Cells(i, J + 1).Value)
                    aDataRow.Add(aVal)
                Next J
                aData.Add(aDataRow)
                i = i + 1
            Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
            Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
            aRetStr = aSAPCostActivityPlanning.PostActivityInput(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData)
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonPostAI_Click")
        End Try
        aDws.Cells(i, 2) = aRetStr
    End Sub

    Private Sub ButtonReadPerAI_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonReadPerAI.Click
        Dim aSAPCOObject As New SAPCOObject
        Dim aCompCodeSplit
        Dim aCompCode
        Dim aData As New Collection
        Dim aContrl As New Collection
        Dim aObjects As New Collection
        Dim aRetStr As String
        Dim i As Integer
        Dim aRange As Excel.Range
        Dim J As Integer
        Dim aVal

        If getParameters("Periodic") = False Then
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
        aRetStr = aSAPCostActivityPlanning.ReadActivityInput(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData)
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
                    aDws.Cells(i + 1, 1).Value = aObjects(i).Costcenter
                    aDws.Cells(i + 1, 2).Value = aObjects(i).WBS_ELEMENT
                    aDws.Cells(i + 1, 3).Value = aObjects(i).Acttype
                    aDws.Cells(i + 1, 4).Value = aObjects(i).SCostcenter
                    aDws.Cells(i + 1, 5).Value = aObjects(i).SActtype
                    aDws.Cells(i + 1, 6).Value = CStr(aSapDataRow.GetValue("UNIT_OF_MEASURE"))
                    For J = 6 To 37
                        aVal = CDbl(aSapDataRow.GetValue(J - 1))
                        aDws.Cells(i + 1, J + 1).Value = aVal
                    Next J
                    i = i + 1
                Loop While i <= aObjects.Count
            End If
            aDws.Cells(i + 1, 2) = aRetStr
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonReadPerAI_Click")
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonPostPerSK_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostPerSK.Click
        ButtonPostSK_Execute("Periodic")
    End Sub

    Private Sub ButtonReadPerSK_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonReadPerSK.Click
        ButtonReadSK_Execute("Periodic")
    End Sub
End Class

