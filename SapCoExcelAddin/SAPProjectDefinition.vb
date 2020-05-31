' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPProjectDefinition
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
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPProjectDefinition")
            log.Error("New - Exception=" & ex.ToString)
        End Try
    End Sub

    Public Function GetPspnr(pPSPID As String) As String
        GetPspnr = ""
        Try
            log.Debug("New - " & "creating Function Z_CO_PS_PROJ_INTERNAL")
            oRfcFunction = destination.Repository.CreateFunction("Z_CO_PS_PROJ_INTERNAL")
            log.Debug("New - " & "oRfcFunction.Metadata.Name=" & oRfcFunction.Metadata.Name)
            RfcSessionManager.BeginContext(destination)
            oRfcFunction.SetValue("I_PSPID", pPSPID)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            GetPspnr = oRfcFunction.GetString("E_PSPNR")
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPProjectDefinition")
            GetPspnr = "Error: Exception in GetPspnr"
            log.Error("GetPspnr - Exception=" & Ex.ToString)
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class

