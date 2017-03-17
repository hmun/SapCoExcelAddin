' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector
Public Class SapCon
    Const aParamWs As String = "Parameter"
    Const aConnectionWs As String = "SAP-Con"
    Private aSapExcelDestinationConfiguration As SapExcelDestinationConfiguration
    Private aDest As String
    Private destination As RfcCustomDestination

    Public Sub New()
        Dim parameters As New RfcConfigParameters()

        Dim aCws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aCws = aWB.Worksheets(aConnectionWs)
        Catch Exc As System.Exception
            MsgBox("No " & aConnectionWs & " Sheet in current workbook", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCon")
            Exit Sub
        End Try
        aDest = aCws.Cells(2, 2).Value
        aSapExcelDestinationConfiguration = New SapExcelDestinationConfiguration
        aSapExcelDestinationConfiguration.ExcelAddOrChangeDestination(aConnectionWs)
        aSapExcelDestinationConfiguration.SetUp()
    End Sub

    Public Function checkCon() As Integer
        Dim dest As RfcDestination
        If destination Is Nothing Then
            Try
                dest = RfcDestinationManager.GetDestination(aDest)
                destination = dest.CreateCustomDestination()
            Catch Ex As System.Exception
                MsgBox("Error reading destination " & aDest & "! Check the connection settings in SAP-Con",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCon")
                checkCon = 16
                Exit Function
            End Try
        End If
        If destination.User = "" Then
            Dim oForm As New FormLogon
            Dim aClient As String
            Dim aUserName As String
            Dim aPassword As String
            Dim aLanguage As String
            Dim aRet As VariantType
            If Not destination.Client Is Nothing Then
                oForm.Client.Text = destination.Client
            End If
            If Not destination.Language Is Nothing Then
                oForm.Language.Text = destination.Language
            End If
            aRet = oForm.ShowDialog()
            If aRet = System.Windows.Forms.DialogResult.OK Then
                aClient = oForm.Client.Text
                aUserName = oForm.UserName.Text
                aPassword = oForm.Password.Text
                aLanguage = oForm.Language.Text
                setCredentials(aClient, aUserName, aPassword, aLanguage)
            End If
        End If
        Try
            destination.Ping()
            checkCon = 0
        Catch ex As RfcInvalidParameterException
            clearCredentials()
            MsgBox("Connecting to SAP failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCon")
            checkCon = 4
        Catch ex As RfcBaseException
            clearCredentials()
            MsgBox("Connecting to SAP failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCon")
            checkCon = 8
        End Try
    End Function

    Public Sub setCredentials(aClient As String, aUsername As String, aPassword As String, aLanguage As String)
        Try
            destination.Client = aClient
            destination.User = aUsername
            destination.Password = aPassword
            destination.Language = aLanguage
        Catch ex As System.Exception
            MsgBox("setCredentials failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCon")
        End Try
    End Sub

    Public Sub SAPlogoff()
        destination = Nothing
        aSapExcelDestinationConfiguration.TearDown()
    End Sub

    Public Sub clearCredentials()
        Try
            destination.User = ""
            destination.Password = Nothing
        Catch ex As System.Exception
            MsgBox("clearCredentials failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCon")
        End Try
    End Sub

    Public Function getDestination() As RfcCustomDestination
        getDestination = destination
    End Function

End Class
