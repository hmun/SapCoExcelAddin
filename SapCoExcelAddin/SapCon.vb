' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector
Public Class SapCon
    Const aParamWs As String = "Parameter"
    Const aConnectionWs As String = "SAP-Con"
    Private aSapExcelDestinationConfiguration As SapExcelDestinationConfiguration
    Private aDest As String
    Public destination As RfcCustomDestination
    Private connected As Boolean = False

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
        aSapExcelDestinationConfiguration.ConfigAddOrChangeDestination()
        aSapExcelDestinationConfiguration.ExcelAddOrChangeDestination(aConnectionWs)
        aSapExcelDestinationConfiguration.SetUp()
        setDest()
    End Sub

    Private Function setDest()
        Dim formRet = 0
        Dim oForm As New FormDestinations
        Dim destCol As Collection
        Dim dest As String
        destCol = aSapExcelDestinationConfiguration.getDestinationList()
        For Each dest In destCol
            oForm.ListBoxDest.Items.Add(dest)
        Next
        formRet = oForm.ShowDialog()
        If formRet = System.Windows.Forms.DialogResult.OK Then
            aDest = oForm.ListBoxDest.SelectedItem.ToString
        Else
            aDest = ""
        End If
    End Function

    Public Function checkCon() As Integer
        Dim dest As RfcDestination = Nothing
        Dim formRet = 0
        If aDest = "" Then
            setDest()
        End If
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
        If Not connected And destination.SncMode = 1 Then
            Dim oForm As New FormLogon
            Dim aClient As String
            Dim aUserName As String
            Dim aPassword As String
            Dim aLanguage As String
            oForm.Destination.Text = dest.Name
            If Not destination.Client Is Nothing Then
                oForm.Client.Text = destination.Client
            End If
            If My.Settings.SAP_Language IsNot Nothing And My.Settings.SAP_Language <> "" Then
                oForm.Language.Text = My.Settings.SAP_Language
            ElseIf Not destination.Language Is Nothing Then
                oForm.Language.Text = destination.Language
            End If
            oForm.UserName.Text = destination.SncMyName
            oForm.UserName.Enabled = False
            oForm.Password.Enabled = False
            formRet = oForm.ShowDialog()
            If formRet = System.Windows.Forms.DialogResult.OK Then
                aClient = oForm.Client.Text
                aUserName = oForm.UserName.Text
                aPassword = oForm.Password.Text
                aLanguage = oForm.Language.Text
                My.Settings.SAP_Language = oForm.Language.Text
                setCredentials_SNC(aClient, aLanguage)
            End If
        ElseIf Not connected Then
            Dim oForm As New FormLogon
            Dim aClient As String
            Dim aUserName As String
            Dim aPassword As String
            Dim aLanguage As String
            Dim aRet As VariantType
            If Not destination.Client Is Nothing Then
                oForm.Client.Text = destination.Client
            End If
            If My.Settings.SAP_Language IsNot Nothing And My.Settings.SAP_Language <> "" Then
                oForm.Language.Text = My.Settings.SAP_Language
            ElseIf Not destination.Language Is Nothing Then
                oForm.Language.Text = destination.Language
            End If
            oForm.Destination.Text = dest.Name
            oForm.UserName.Enabled = True
            If My.Settings.SAP_User IsNot Nothing Then
                oForm.UserName.Text = CStr(My.Settings.SAP_User)
            End If
            oForm.Password.Enabled = True
            formRet = oForm.ShowDialog()
            If formRet = System.Windows.Forms.DialogResult.OK Then
                aClient = oForm.Client.Text
                aUserName = oForm.UserName.Text
                My.Settings.SAP_User = oForm.UserName.Text
                aPassword = oForm.Password.Text
                aLanguage = oForm.Language.Text
                My.Settings.SAP_Language = oForm.Language.Text
                setCredentials(aClient, aUserName, aPassword, aLanguage)
            End If
        End If
        If connected Or formRet = System.Windows.Forms.DialogResult.OK Then
            Try
                destination.Ping()
                connected = True
                checkCon = 0
            Catch ex As RfcInvalidParameterException
                clearCredentials()
                MsgBox("Connecting to SAP failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCon")
                connected = False
                checkCon = 4
            Catch ex As RfcBaseException
                clearCredentials()
                MsgBox("Connecting to SAP failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCon")
                connected = False
                checkCon = 8
            End Try
        Else
            connected = False
            destination = Nothing
            checkCon = 8
        End If
    End Function

    Public Sub setCredentials_SNC(aClient As String, aLanguage As String)
        Try
            destination.Client = aClient
            destination.Language = aLanguage
        Catch ex As System.Exception
            MsgBox("setCredentials failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCon")
        End Try
    End Sub

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
        If aDest IsNot Nothing And aDest <> "" Then
            aSapExcelDestinationConfiguration.TearDown(aDest)
        Else
            aSapExcelDestinationConfiguration.TearDown()
        End If
        connected = False
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
