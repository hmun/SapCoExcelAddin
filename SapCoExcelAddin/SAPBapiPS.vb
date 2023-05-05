' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/
Imports SAP.Middleware.Connector

Public Class SAPBapiPS

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        Try
            sapcon = aSapCon
            aSapCon.getDestination(destination)
            sapcon.checkCon()
        Catch ex As System.Exception
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPBapiPS")
        End Try
    End Sub

    Public Function initialization() As Integer
        log.Debug("New - " & "creating Function BAPI_PS_INITIALIZATION")
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_PS_INITIALIZATION")
            log.Debug("initialization - " & "oRfcFunction.Metadata.Name=" & oRfcFunction.Metadata.Name)
            oRfcFunction.Invoke(destination)
            initialization = True
        Catch Exc As System.Exception
            log.Error("initialization - Exception=" & Exc.ToString)
            initialization = False
        End Try
    End Function

    Public Function precommit() As String
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_PS_PRECOMMIT")
            log.Debug("initialization - " & "oRfcFunction.Metadata.Name=" & oRfcFunction.Metadata.Name)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            oRETURN.Clear()
            oRfcFunction.Invoke(destination)
            precommit = ""
            Dim aErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
                precommit = precommit & ";" & oRETURN(i).GetValue("MESSAGE")
            Next i
            precommit = If(aErr = True, "Error: " & precommit, precommit)
        Catch Exc As System.Exception
            log.Error("initialization - Exception=" & Exc.ToString)
            precommit = False
        End Try
    End Function

End Class
