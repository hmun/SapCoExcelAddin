' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPProjectDefinition
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        Try
            sapcon = aSapCon
            destination = aSapCon.getDestination()
            sapcon.checkCon()
        Catch ex As System.Exception
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPProjectDefinition")
        End Try
    End Sub

    Public Function GetPspnr(pPSPID As String) As String
        GetPspnr = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("Z_CO_PS_PROJ_INTERNAL")
            RfcSessionManager.BeginContext(destination)
            oRfcFunction.SetValue("I_PSPID", pPSPID)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            GetPspnr = oRfcFunction.GetString("E_PSPNR")
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPProjectDefinition")
            GetPspnr = "Error: Exception in GetPspnr"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class

