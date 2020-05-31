' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPWbsElement

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        sapcon = aSapCon
        aSapCon.getDestination(destination)
        log.Debug("New - " & "creating Function Z_CO_PS_PSP_INTERNAL")
        Try
            oRfcFunction = destination.Repository.CreateFunction("Z_CO_PS_PSP_INTERNAL")
            log.Debug("New - " & "oRfcFunction.Metadata.Name=" & oRfcFunction.Metadata.Name)
        Catch ex As Exception
            oRfcFunction = Nothing
            log.Error("New - Exception=" & ex.ToString)
        End Try
    End Sub

    Public Function GetPspnr(pPOSID As String) As String
        If Not oRfcFunction Is Nothing Then
            sapcon.checkCon()
            Try
                log.Debug("GetPspnr - " & "Setting Function parameters")
                oRfcFunction.SetValue("I_POSID", pPOSID)
                log.Debug("GetPspnr - " & "invoking " & oRfcFunction.Metadata.Name)
                oRfcFunction.Invoke(destination)
                GetPspnr = oRfcFunction.GetValue("E_PSPNR")
                Exit Function
            Catch ex As Exception
                MsgBox("Exception in GetPspnr! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWbsElement")
                log.Error("GetPspnr - Exception=" & ex.ToString)
                GetPspnr = "Fehler"
            End Try
        Else
            GetPspnr = pPOSID
        End If
    End Function

End Class
