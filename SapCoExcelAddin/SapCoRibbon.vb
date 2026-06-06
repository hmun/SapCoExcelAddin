' Copyright 2026 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports Microsoft.Office.Tools.Ribbon
Imports System.Configuration
Imports System.Collections.Specialized

Public Class SapCoRibbon
    Private aSapCon
    Private aSapGeneral
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
            log.Debug("No Parameter_Int Sheet in current workbook. - ignored")
            getIntParameters = True
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
            MsgBox("SAP-Logon successful! ", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Sap Co")
        Else
            log.Debug("ButtonLogon_Click - " & "connection failed")
            aSapCon = Nothing
        End If
    End Sub

    Private Sub SapCoRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        Dim sAll As NameValueCollection
        Dim s As String
        Dim enableGeneration As Boolean = False
        Dim enableGenerationWBS As Boolean = False
        Dim enablePS As Boolean = False
        Dim enablePSGetWbs As Boolean = False
        aSapGeneral = New SapGeneral
        Try
            sAll = ConfigurationManager.AppSettings
            s = sAll("enableGeneration")
            enableGeneration = Convert.ToBoolean(s)
            s = sAll("enableGenerationWBS")
            enableGenerationWBS = Convert.ToBoolean(s)
            s = sAll("enablePS")
            enablePS = Convert.ToBoolean(s)
            s = sAll("enablePSGetWbs")
            enablePSGetWbs = Convert.ToBoolean(s)

        Catch Exc As System.Exception
            log.Error("SapCoRibbon_Load - " & "Exception=" & Exc.ToString)
        End Try
        If Not enableGeneration Then
            Globals.Ribbons.SapCoRibbon.SAP_COGenerate.Visible = False
        Else
            Globals.Ribbons.SapCoRibbon.SAP_COGenerate.Visible = True
        End If
        If Not enableGenerationWBS Then
            Globals.Ribbons.SapCoRibbon.ButtonGenerateWbs.Visible = False
        Else
            Globals.Ribbons.SapCoRibbon.ButtonGenerateWbs.Visible = True
        End If
        If Not enablePS Then
            Globals.Ribbons.SapCoRibbon.SAP_COWbs.Visible = False
        Else
            Globals.Ribbons.SapCoRibbon.SAP_COWbs.Visible = True
        End If
        If Not enablePSGetWbs Then
            Globals.Ribbons.SapCoRibbon.ButtonGetWbs.Visible = False
        Else
            Globals.Ribbons.SapCoRibbon.ButtonGetWbs.Visible = True
        End If
    End Sub


    Private Sub ButtonActivityAllocCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonActivityAllocCheck.Click
        Dim aSapClass As New SapCoRibbon_AcctngActivityAlloc
        If checkCon() = True Then
            aSapClass.Post(pSapCon:=aSapCon, pCheck:=True)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonActivityAllocCheck")
        End If
    End Sub

    Private Sub ButtonActivityAllocPost_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonActivityAllocPost.Click
        Dim aSapClass As New SapCoRibbon_AcctngActivityAlloc
        If checkCon() = True Then
            aSapClass.Post(pSapCon:=aSapCon, pCheck:=False)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonActivityAllocPost")
        End If
    End Sub

    Private Sub ButtonRepstPrimCostsCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonRepstPrimCostsCheck.Click
        Dim aSapClass As New SapCoRibbon_AcctngRepstPrimCosts
        If checkCon() = True Then
            aSapClass.Post(pSapCon:=aSapCon, pCheck:=True)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonRepstPrimCostsCheck")
        End If
    End Sub

    Private Sub ButtonRepstPrimCostsPost_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonRepstPrimCostsPost.Click
        Dim aSapClass As New SapCoRibbon_AcctngRepstPrimCosts
        If checkCon() = True Then
            aSapClass.Post(pSapCon:=aSapCon, pCheck:=False)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonRepstPrimCostsPost")
        End If
    End Sub

    Private Sub ButtonManCostAllocCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonManCostAllocCheck.Click
        Dim aSapClass As New SapCoRibbon_AcctngManCostAlloc
        If checkCon() = True Then
            aSapClass.Post(pSapCon:=aSapCon, pCheck:=True)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonManCostAllocCheck")
        End If
    End Sub

    Private Sub ButtonManCostAllocPost_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonManCostAllocPost.Click
        Dim aSapClass As New SapCoRibbon_AcctngManCostAlloc
        If checkCon() = True Then
            aSapClass.Post(pSapCon:=aSapCon, pCheck:=False)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonManCostAllocPost")
        End If
    End Sub

    Private Sub ButtonCheckCostingBasedData_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCheckCostingBasedData.Click
        Dim aSapClass As New SapCoRibbon_PostCOPA
        If checkCon() = True Then
            aSapClass.Post(pSapCon:=aSapCon, pCheck:=True)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonCheckCostingBasedData")
        End If
    End Sub

    Private Sub ButtonPostCostingBasedData_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostCostingBasedData.Click
        Dim aSapClass As New SapCoRibbon_PostCOPA
        If checkCon() = True Then
            aSapClass.Post(pSapCon:=aSapCon, pCheck:=False)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonPostCostingBasedData")
        End If
    End Sub


    Private Sub ButtonStatKeyFiguresCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonStatKeyFiguresCheck.Click
        Dim aSapClass As New SapCoRibbon_AcctngStatKeyFigures
        If checkCon() = True Then
            aSapClass.Post(pSapCon:=aSapCon, pCheck:=True)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonStatKeyFiguresCheck")
        End If
    End Sub

    Private Sub ButtonStatKeyFiguresPost_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonStatKeyFiguresPost.Click
        Dim aSapClass As New SapCoRibbon_AcctngStatKeyFigures
        If checkCon() = True Then
            aSapClass.Post(pSapCon:=aSapCon, pCheck:=False)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonStatKeyFiguresPost")
        End If
    End Sub

    Private Sub ButtonGenerate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonGenerate.Click
        Dim aRibbon_Generate As New Ribbon_Generate(pSapGeneral:=aSapGeneral)
        aRibbon_Generate.GenerateData()
    End Sub

    Private Sub ButtonGenerateWbs_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonGenerateWbs.Click
        Dim aRibbon_Generate As New Ribbon_Generate(pSapGeneral:=aSapGeneral)
        aRibbon_Generate.GenerateData()
    End Sub

    Private Sub ButtonCreateWbs_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCreateWbs.Click
        Dim aSapPsMdRibbonWbs As New SapPsMdRibbonWbs
        If checkCon() = True Then
            aSapPsMdRibbonWbs.exec(pSapCon:=aSapCon, pMode:="Create")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonGetWbs_Click")
        End If
    End Sub

    Private Sub ButtonGetWbs_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonGetWbs.Click
        Dim aSapPsMdRibbonWbs As New SapPsMdRibbonWbs
        If checkCon() = True Then
            aSapPsMdRibbonWbs.GetData(pSapCon:=aSapCon)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonGetWbs_Click")
        End If
    End Sub

    Private Sub ButtonWBSSetStatus_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonWBSSetStatus.Click
        Dim aSapPsMdRibbonWbs As New SapPsMdRibbonWbs
        If checkCon() = True Then
            aSapPsMdRibbonWbs.Status(pSapCon:=aSapCon, pMode:="Set")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonWBSSetStatus_Click")
        End If
    End Sub

End Class