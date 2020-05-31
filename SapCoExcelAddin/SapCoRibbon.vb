' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports Microsoft.Office.Tools.Ribbon
Imports SAP.Middleware.Connector

Public Class SapCoRibbon
    Private aSapCon
    Private aSapGeneral
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Const AI_CM = 22 ' column of message for activity posting
    Const PC_CM = 21 ' column of message for primary cost reposting
    Const MC_CM = 20 ' column of message for manual cost allocation

    Private aCoAre As String
    Private aOperatingConcern As String
    Private aMaxLines As String
    Private aIgnoreSelf As String
    Private aFromLine As String
    Private aToLine As String
    Private aDocDate As String
    Private aPostDate As String

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

    Private Sub ButtonLogoff_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogoff.Click
        log.Debug("ButtonLogoff_Click - " & "starting logoff")
        If Not aSapCon Is Nothing Then
            log.Debug("ButtonLogoff_Click - " & "calling aSapCon.SAPlogoff()")
            aSapCon.SAPlogoff()
            aSapCon = Nothing
        End If
        log.Debug("ButtonLogoff_Click - " & "exit")
    End Sub

    Private Sub ButtonLogon_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogon.Click
        Dim aConRet As Integer

        log.Debug("ButtonLogon_Click - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            log.Debug("ButtonLogon_Click - " & "Version check failed")
            Exit Sub
        End If
        log.Debug("ButtonLogon_Click - " & "creating SapCon")
        If aSapCon Is Nothing Then
            aSapCon = New SapCon()
        End If
        log.Debug("ButtonLogon_Click - " & "calling SapCon.checkCon()")
        aConRet = aSapCon.checkCon()
        If aConRet = 0 Then
            log.Debug("ButtonLogon_Click - " & "connection successfull")
            MsgBox("SAP-Logon successful! ", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Sap Accounting")
        Else
            log.Debug("ButtonLogon_Click - " & "connection failed")
            aSapCon = Nothing
        End If
    End Sub

    Private Sub SapCoRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        aSapGeneral = New SapGeneral
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
        Dim i As Integer
        i = 4
        Do While CStr(aPws.Cells(i, 1).Value) <> ""
            If CStr(aPws.Cells(i, 1).Value) = "IgnoreSelf" Then
                aIgnoreSelf = CStr(aPws.Cells(i, 2).Value)
            End If
            If CStr(aPws.Cells(i, 1).Value) = "FromLine" Then
                aFromLine = CStr(aPws.Cells(i, 2).Value)
            End If
            If CStr(aPws.Cells(i, 1).Value) = "ToLine" Then
                aToLine = CStr(aPws.Cells(i, 2).Value)
            End If
            i = i + 1
        Loop
        If aIgnoreSelf Is Nothing Then
            aIgnoreSelf = ""
        End If
        If aFromLine Is Nothing Then
            aFromLine = ""
        End If
        If aToLine Is Nothing Then
            aToLine = ""
        End If
        getActivityAllocParameters = True
    End Function

    Private Sub ButtonActivityAllocCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonActivityAllocCheck.Click
        If checkCon() = True Then
            SAP_ActivityAlloc_execute(pTest:=True)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonActivityAllocCheck_Click")
        End If
    End Sub

    Private Sub ButtonActivityAllocPost_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonActivityAllocPost.Click
        If checkCon() = True Then
            SAP_ActivityAlloc_execute(pTest:=False)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonActivityAllocPost_Click")
        End If
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
        Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
        Globals.SapCoExcelAddin.Application.EnableEvents = False
        Globals.SapCoExcelAddin.Application.ScreenUpdating = False
        Dim aBUDAT As String
        Dim aBLDAT As String
        Dim aCells As Excel.Range
        Dim aFrom As UInteger
        Dim aTo As UInteger
        aBUDAT = ""
        aBLDAT = ""
        If aFromLine <> "" And aToLine <> "" Then
            aFrom = CUInt(aFromLine)
            aTo = CUInt(aToLine)
        Else
            aFrom = 2
            aTo = UInt32.MaxValue
        End If
        Try
            i = aFrom
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
                                                                           CStr(aDws.Cells(i, 5).Value),
                                                                           CDbl(FormatNumber(CDbl(aDws.Cells(i, 6).Value), 3, True, False, False)),
                                                                           CStr(aDws.Cells(i, 7).Value), CStr(aDws.Cells(i, 8).Value),
                                                                           CStr(aDws.Cells(i, 9).Value), CStr(aDws.Cells(i, 10).Value),
                                                                           CStr(aDws.Cells(i, 11).Value), CStr(aDws.Cells(i, 12).Value),
                                                                           CDbl(FormatNumber(CDbl(aDws.Cells(i, 14).Value), 2, True, False, False)),
                                                                           CDbl(FormatNumber(CDbl(aDws.Cells(i, 15).Value), 2, True, False, False)),
                                                                           CDbl(FormatNumber(CDbl(aDws.Cells(i, 16).Value), 2, True, False, False)),
                                                                           CInt(aDws.Cells(i, 17).Value),
                                                                           CStr(aDws.Cells(i, 21).Value), CStr(aDws.Cells(i, 13).Value),
                                                                           CDbl(FormatNumber(CDbl(aDws.Cells(i, 18).Value), 2, True, False, False)),
                                                                           CDbl(FormatNumber(CDbl(aDws.Cells(i, 19).Value), 2, True, False, False)),
                                                                           CDbl(FormatNumber(CDbl(aDws.Cells(i, 20).Value), 2, True, False, False)))
                    If aIgnoreSelf.ToUpper() = "" Or aSAPAcctngActivityItem.SEND_CCTR <> aSAPAcctngActivityItem.REC_CCTR Then
                        aData.Add(aSAPAcctngActivityItem)
                        If aLines >= CInt(aMaxLines) Then
                            aRetStr = aSAPAcctngActivityAlloc.post(aCoAre, CDate(aBUDAT), CDate(aBLDAT), aData, pTest)
                            aCells = aDws.Range(aDws.Cells(aPostLine + 1, AI_CM), aDws.Cells(i, AI_CM))
                            aCells.Value = aRetStr
                            aData = New Collection
                            aRetStr = ""
                            aLines = 1
                            aBUDAT = ""
                            aPostLine = i
                        Else
                            aLines += 1
                        End If
                    Else
                        If aPostLine = i - 1 Then
                            aPostLine += 1
                        End If
                    End If
                End If
                i += 1
            Loop While CStr(aDws.Cells(i, 1).value) <> "" And i <= aTo
            If aData.Count > 0 Then
                aRetStr = aSAPAcctngActivityAlloc.post(aCoAre, CDate(aBUDAT), CDate(aBLDAT), aData, pTest)
                aCells = aDws.Range(aDws.Cells(aPostLine + 1, AI_CM), aDws.Cells(i - 1, AI_CM))
                aCells.Value = aRetStr
                aRetStr = ""
            End If
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch Ex As System.Exception
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
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
            MsgBox("Cell A1 of the parameter sheet does not contain the key SAPAcctngRepstPrimCosts. Check if the current workbook is a valid SAP CO RepstPrimCosts Template",
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
        If checkCon() = True Then
            SAP_RepstPrimCosts_execute(pTest:=True)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonRepstPrimCostsCheck_Click")
        End If
    End Sub

    Private Sub ButtonRepstPrimCostsPost_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonRepstPrimCostsPost.Click
        If checkCon() = True Then
            SAP_RepstPrimCosts_execute(pTest:=False)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonRepstPrimCostsPost_Click")
        End If
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
        Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
        Globals.SapCoExcelAddin.Application.EnableEvents = False
        Globals.SapCoExcelAddin.Application.ScreenUpdating = False
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
                                                                             CStr(aDws.Cells(i, 11).Value), CDbl(FormatNumber(aDws.Cells(i, 12).Value, 2, True, False, False)),
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
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch Ex As System.Exception
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_RepstPrimCosts_execute")
        End Try
    End Sub

    Private Function getManCostAllocParameters() As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP CO ManCostAlloc Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getManCostAllocParameters = False
            Exit Function
        End Try
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> "SAPAcctngManCostAlloc" Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key SAPManCostAlloc. Check if the current workbook is a valid SAP CO ManCostAlloc Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getManCostAllocParameters = False
            Exit Function
        End If
        aCoAre = CStr(aPws.Cells(2, 2).Value)
        aMaxLines = CInt(aPws.Cells(3, 2).Value)
        If aCoAre = "" Then
            MsgBox("Please fill all obligatory fields in the parameter sheet!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ManCostAlloc")
            getManCostAllocParameters = False
            Exit Function
        End If
        getManCostAllocParameters = True
    End Function

    Private Sub ButtonManCostAllocCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonManCostAllocCheck.Click
        If checkCon() = True Then
            SAP_ManCostAlloc_execute(pTest:=True)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonManCostAllocCheck_Click")
        End If
    End Sub

    Private Sub ButtonManCostAllocPost_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonManCostAllocPost.Click
        If checkCon() = True Then
            SAP_ManCostAlloc_execute(pTest:=False)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonManCostAllocPost_Click")
        End If
    End Sub

    Private Sub SAP_ManCostAlloc_execute(pTest As Boolean)
        Dim i As Integer
        Dim aLines As Integer
        Dim aPostLine As Integer
        Dim aData As New Collection
        Dim aRetStr As String
        Dim aDateFormatString As New DateFormatString
        Dim aSAPAcctngPrimCostsItem As New SAPAcctngPrimCostsItem

        If getManCostAllocParameters() = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        Dim aSAPAcctngManCostAlloc As New SAPAcctngManCostAlloc(aSapCon)
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("Data")
        Catch Exc As System.Exception
            MsgBox("No Data Sheet in current workbook. Check if the current workbook is a valid SAP CO ManCostAlloc Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            Exit Sub
        End Try
        aRetStr = ""
        aDws.Activate()
        Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
        Globals.SapCoExcelAddin.Application.EnableEvents = False
        Globals.SapCoExcelAddin.Application.ScreenUpdating = False
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
                If InStr(CStr(aDws.Cells(i, MC_CM).Value), "Beleg wird unter der Nummer") = 0 And
                   InStr(CStr(aDws.Cells(i, MC_CM).Value), "Document is posted under number") = 0 Then
                    If aBUDAT = "" Or aMaxLines = 1 Then
                        aBUDAT = Format(aDws.Cells(i, 1).Value, aDateFormatString.getString)
                        aBLDAT = Format(aDws.Cells(i, 2).Value, aDateFormatString.getString)
                    End If
                    aSAPAcctngPrimCostsItem = aSAPAcctngPrimCostsItem.create(CStr(aDws.Cells(i, 3).Value), "", CStr(aDws.Cells(i, 4).Value),
                                                                             CStr(aDws.Cells(i, 5).Value), CStr(aDws.Cells(i, 6).Value),
                                                                             CStr(aDws.Cells(i, 7).Value), CStr(aDws.Cells(i, 8).Value),
                                                                             CStr(aDws.Cells(i, 9).Value), CStr(aDws.Cells(i, 10).Value),
                                                                             CDbl(FormatNumber(aDws.Cells(i, 11).Value, 2, True, False, False)),
                                                                             CStr(aDws.Cells(i, 12).Value),
                                                                             CStr(aDws.Cells(i, 13).Value), CStr(aDws.Cells(i, 14).Value),
                                                                             CStr(aDws.Cells(i, 15).Value), CStr(aDws.Cells(i, 16).Value),
                                                                             CStr(aDws.Cells(i, 17).Value), CStr(aDws.Cells(i, 18).Value),
                                                                             CStr(aDws.Cells(i, 19).Value))
                    aData.Add(aSAPAcctngPrimCostsItem)
                    If aLines >= CInt(aMaxLines) Then
                        aRetStr = aSAPAcctngManCostAlloc.post(aCoAre, CDate(aBUDAT), CDate(aBLDAT), aData, pTest)
                        aCells = aDws.Range(aDws.Cells(aPostLine + 1, MC_CM), aDws.Cells(i, MC_CM))
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
                aRetStr = aSAPAcctngManCostAlloc.post(aCoAre, CDate(aBUDAT), CDate(aBLDAT), aData, pTest)
                aCells = aDws.Range(aDws.Cells(aPostLine + 1, MC_CM), aDws.Cells(i - 1, MC_CM))
                aCells.Value = aRetStr
            End If
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch Ex As System.Exception
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_ManCostAlloc_execute")
        End Try
    End Sub

    Private Function getCOPAParameters() As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP CO-PA Actuals Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getCOPAParameters = False
            Exit Function
        End Try
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> "SAPCostingBasedData" Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key SAPCostingBasedData. Check if the current workbook is a valid SAP CO-PA Actuals Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getCOPAParameters = False
            Exit Function
        End If
        aOperatingConcern = CStr(aPws.Cells(2, 2).Value)
        aMaxLines = CInt(aPws.Cells(3, 2).Value)
        If aOperatingConcern = "" Then
            MsgBox("Please fill all obligatory fields in the parameter sheet!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap RepstPrimCosts")
            getCOPAParameters = False
            Exit Function
        End If
        getCOPAParameters = True
    End Function

    Private Sub ButtonCheckCostingBasedData_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCheckCostingBasedData.Click
        If checkCon() = True Then
            SAP_COPA_exec(pTest:=True)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonPostCostingBasedData_Click")
        End If
    End Sub

    Private Sub ButtonPostCostingBasedData_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostCostingBasedData.Click
        If checkCon() = True Then
            SAP_COPA_exec(pTest:=False)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonPostCostingBasedData_Click")
        End If
    End Sub

    Private Sub SAP_COPA_exec(pTest As Boolean)
        Dim aSAPCOPAActuals As New SAPCOPAActuals(aSapCon)
        Dim aSAPCOPAItem As New SAPCOPAItem
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        Dim aSAPFormat As New SAPFormat
        Dim aSAPProjectDefinition As New SAPProjectDefinition(aSapCon)
        Dim aSAPWbsElement As New SAPWbsElement(aSapCon)
        Dim aData As New Collection
        Dim aDataRow As New Collection
        Dim aLines As Integer
        Dim aStartLine As Integer
        Dim aEndLine As Integer
        Dim aLineCnt As Integer

        Dim i As Integer
        Dim j As Integer
        Dim maxJ As Integer
        Dim aRetStr As String

        Dim aFIELDNAME As String
        Dim aVALUE As Object
        Dim aCURRENCY As String

        Dim aCells As Excel.Range

        If getCOPAParameters() = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("Data")
        Catch Exc As System.Exception
            MsgBox("No Data Sheet in current workbook. Check if the current workbook is a valid SAP CO-PA Actuals Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-PA")
            Exit Sub
        End Try
        ' Read the Items
        aDws.Activate()
        Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
        Globals.SapCoExcelAddin.Application.EnableEvents = False
        Globals.SapCoExcelAddin.Application.ScreenUpdating = False
        i = 5
        ' determine the last column
        maxJ = 1
        Do
            maxJ = maxJ + 1
        Loop While CStr(aDws.Cells(1, maxJ).Value) <> ""
        aStartLine = i
        aLineCnt = 0
        aData = New Collection
        Do
            If Left(CStr(aDws.Cells(i, maxJ).Value), 7) <> "Success" Then
                aDataRow = New Collection
                j = 1
                Do
                    aVALUE = ""
                    aCURRENCY = ""
                    aFIELDNAME = ""
                    aSAPCOPAItem = New SAPCOPAItem
                    If aDws.Cells(2, j).Value IsNot Nothing Then
                        aCURRENCY = CStr(aDws.Cells(2, j).Value)
                        If aDws.Cells(i, j).Value IsNot Nothing Then
                            aVALUE = FormatNumber(CDbl(aDws.Cells(i, j).Value), 2, True, False, False)
                        Else
                            aVALUE = FormatNumber(0, 2, True, False, False)
                        End If
                    Else
                        aCURRENCY = ""
                        If aDws.Cells(i, j).Value IsNot Nothing Then
                            Select Case CStr(aDws.Cells(3, j).Value)
                                Case "DATE"
                                    Try
                                        aVALUE = CDate(aDws.Cells(i, j).Value).ToString("yyyyMMdd")
                                    Catch Exc As System.Exception
                                        aVALUE = ""
                                    End Try
                                Case "PERIO"
                                    aVALUE = Right(aDws.Cells(i, j).Value, 4) & Left(aDws.Cells(i, j).Value, 3)
                                Case "PROJ"
                                    If CStr(aDws.Cells(i, j).Value) <> "" Then
                                        aVALUE = aSAPProjectDefinition.GetPspnr(CStr(aDws.Cells(i, j).Value))
                                    Else
                                        aVALUE = ""
                                    End If
                                Case "WBS"
                                    If CStr(aDws.Cells(i, j).Value) <> "" Then
                                        aVALUE = aSAPWbsElement.GetPspnr(CStr(aDws.Cells(i, j).Value))
                                    Else
                                        aVALUE = ""
                                    End If
                                Case Else
                                    If Left(aDws.Cells(3, j).Value, 1) = "U" Then
                                        aVALUE = aSAPFormat.unpack(aDws.Cells(i, j).Value, CInt(Right(aDws.Cells(3, j).Value, Len(aDws.Cells(3, j).Value) - 1)))
                                    ElseIf Left(aDws.Cells(3, j).Value, 1) = "P" Then
                                        aVALUE = aSAPFormat.pspid(aDws.Cells(i, j).Value, CInt(Right(aDws.Cells(3, j).Value, Len(aDws.Cells(3, j).Value) - 1)))
                                    Else
                                        aVALUE = aDws.Cells(i, j).Value
                                    End If
                            End Select
                        End If
                    End If

                    aFIELDNAME = CStr(aDws.Cells(1, j).Value)
                    aSAPCOPAItem = aSAPCOPAItem.create(aFIELDNAME, aVALUE, aCURRENCY)
                    aDataRow.Add(aSAPCOPAItem)
                    j = j + 1
                Loop While CStr(aDws.Cells(1, j).Value) <> ""
                aData.Add(aDataRow)
                aLineCnt = aLineCnt + 1
                If aLineCnt >= CInt(aMaxLines) Then
                    aEndLine = i
                    '     post the lines
                    Globals.SapCoExcelAddin.Application.StatusBar = "Posting at line " & aEndLine
                    aRetStr = aSAPCOPAActuals.PostCostingBasedData(aOperatingConcern, aData, pCheck:=pTest)
                    aCells = aDws.Range(aDws.Cells(aStartLine, j), aDws.Cells(aEndLine, j))
                    aCells.Value = aRetStr
                    aStartLine = i + 1
                    aLineCnt = 0
                    aData = New Collection
                End If
            Else
                aDws.Cells(i, maxJ + 1).Value = "ignored - already posted"
            End If
            i = i + 1
        Loop While CStr(aDws.Cells(i, 1).Value) <> ""
        ' post the rest
        If aData.Count > 0 Then
            aEndLine = i - 1
            Globals.SapCoExcelAddin.Application.StatusBar = "Posting at line " & aEndLine
            aRetStr = aSAPCOPAActuals.PostCostingBasedData(aOperatingConcern, aData, pCheck:=pTest)
            aCells = aDws.Range(aDws.Cells(aStartLine, j), aDws.Cells(aEndLine, j))
            aCells.Value = aRetStr
        End If
        Globals.SapCoExcelAddin.Application.EnableEvents = True
        Globals.SapCoExcelAddin.Application.ScreenUpdating = True
        Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
    End Sub

    Private Function getStatKeyFiguresParameters() As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim aDateFormatString As New DateFormatString
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP CO StatKeyFigure Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getStatKeyFiguresParameters = False
            Exit Function
        End Try
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> "SAPAcctngPostStatKeyFigure" Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key SAPAcctngPostStatKeyFigure. Check if the current workbook is a valid SAP CO StatKeyFigure Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getStatKeyFiguresParameters = False
            Exit Function
        End If
        aPostDate = Format(aPws.Cells(2, 2).Value, aDateFormatString.getString)
        aDocDate = Format(aPws.Cells(3, 2).Value, aDateFormatString.getString)
        aCoAre = CStr(aPws.Cells(4, 2).Value)
        aMaxLines = CInt(aPws.Cells(5, 2).Value)
        If aCoAre = "" Then
            MsgBox("Please fill all obligatory fields in the parameter sheet!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap StatKeyFigure")
            getStatKeyFiguresParameters = False
            Exit Function
        End If
        getStatKeyFiguresParameters = True
    End Function

    Private Sub ButtonStatKeyFiguresCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonStatKeyFiguresCheck.Click
        If checkCon() = True Then
            SAP_StatKeyFigures_execute(pTest:=True)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonStatKeyFiguresCheck_Click")
        End If
    End Sub

    Private Sub ButtonStatKeyFiguresPost_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonStatKeyFiguresPost.Click
        If checkCon() = True Then
            SAP_StatKeyFigures_execute(pTest:=False)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonStatKeyFiguresPost_Click")
        End If
    End Sub

    Private Sub SAP_StatKeyFigures_execute(pTest As Boolean)
        Dim i As Integer
        Dim j As Integer
        Dim maxJ As Integer
        Dim aLines As Integer
        Dim aPostLine As Integer
        Dim aData As New Collection
        Dim aRetStr As String
        Dim aDateFormatString As New DateFormatString
        Dim aSapAcctngStatKeyFiguresDocItem As New SapAcctngStatKeyFiguresDocItem
        Dim aSAPFormat As New SAPFormat
        Dim aFIELDNAME As String
        Dim aVALUE As Object

        If getStatKeyFiguresParameters() = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        Dim aSAPAcctngStatKeyFigures As New SapAcctngStatKeyFigures(aSapCon)
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("Data")
        Catch Exc As System.Exception
            MsgBox("No Data Sheet in current workbook. Check if the current workbook is a valid SAP CO StatKeyFigures Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            Exit Sub
        End Try
        aRetStr = ""
        aDws.Activate()
        Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
        Globals.SapCoExcelAddin.Application.EnableEvents = False
        Globals.SapCoExcelAddin.Application.ScreenUpdating = False
        ' determine the last column
        maxJ = 1
        Do
            maxJ += 1
        Loop While CStr(aDws.Cells(1, maxJ).Value) <> ""

        Dim aCells As Excel.Range
        Try
            i = 4
            aLines = 1
            aPostLine = i - 1
            Do
                If InStr(CStr(aDws.Cells(i, maxJ).Value), "Beleg wird unter der Nummer") = 0 And
                   InStr(CStr(aDws.Cells(i, maxJ).Value), "Document is posted under number") = 0 Then
                    ' fill the items here
                    aSapAcctngStatKeyFiguresDocItem = New SapAcctngStatKeyFiguresDocItem
                    j = 1
                    Do
                        aVALUE = ""
                        aFIELDNAME = ""
                        If aDws.Cells(i, j).Value IsNot Nothing Then
                            aFIELDNAME = CStr(aDws.Cells(1, j).Value)
                            Select Case CStr(aDws.Cells(2, j).Value)
                                Case "DOUBLE"
                                    aVALUE = FormatNumber(CDbl(aDws.Cells(i, j).Value), 3, True, False, False)
                                    aSapAcctngStatKeyFiguresDocItem.SetField(aFIELDNAME, aVALUE, "F")
                                Case "DATE"
                                    Try
                                        aVALUE = CDate(aDws.Cells(i, j).Value).ToString("yyyyMMdd")
                                    Catch Exc As System.Exception
                                        aVALUE = ""
                                    End Try
                                    aSapAcctngStatKeyFiguresDocItem.SetField(aFIELDNAME, aVALUE, "S")
                                Case Else
                                    If Left(aDws.Cells(2, j).Value, 1) = "U" Then
                                        aVALUE = aSAPFormat.unpack(aDws.Cells(i, j).Value, CInt(Right(aDws.Cells(2, j).Value, Len(aDws.Cells(2, j).Value) - 1)))
                                    ElseIf Left(aDws.Cells(2, j).Value, 1) = "P" Then
                                        aVALUE = aSAPFormat.pspid(aDws.Cells(i, j).Value, CInt(Right(aDws.Cells(2, j).Value, Len(aDws.Cells(2, j).Value) - 1)))
                                    Else
                                        aVALUE = aDws.Cells(i, j).Value
                                    End If
                                    aSapAcctngStatKeyFiguresDocItem.SetField(aFIELDNAME, aVALUE, "S")
                            End Select
                        End If
                        j += 1
                    Loop While CStr(aDws.Cells(1, j).Value) <> ""
                    aData.Add(aSapAcctngStatKeyFiguresDocItem)
                    If aLines >= CInt(aMaxLines) Then
                        aRetStr = aSAPAcctngStatKeyFigures.post(aCoAre, CDate(aPostDate), CDate(aDocDate), aData, pTest)
                        aCells = aDws.Range(aDws.Cells(aPostLine + 1, maxJ), aDws.Cells(i, maxJ))
                        aCells.Value = aRetStr
                        aData = New Collection
                        aLines = 1
                        aPostLine = i
                    Else
                        aLines += 1
                    End If
                End If
                i += 1
            Loop While CStr(aDws.Cells(i, 1).value) <> ""
            If aData.Count > 0 Then
                aRetStr = aSAPAcctngStatKeyFigures.post(aCoAre, CDate(aPostDate), CDate(aDocDate), aData, pTest)
                aCells = aDws.Range(aDws.Cells(aPostLine + 1, maxJ), aDws.Cells(i - 1, maxJ))
                aCells.Value = aRetStr
            End If
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch Ex As System.Exception
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP_StatKeyFigures_execute")
        End Try
    End Sub

End Class