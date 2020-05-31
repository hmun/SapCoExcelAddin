' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector
Public Class SapAcctngStatKeyFigures
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
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAcctngRepstPrimCosts")
        End Try
    End Sub

    Public Function post(pKokrs As String, pBuDat As Date, pBldat As Date, pData As Collection, pTest As Boolean) As String
        post = ""
        Try
            If pTest Then
                oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_STAT_KEY_FIG_CHECK")
            Else
                oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_STAT_KEY_FIG_POST")
            End If
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oDocHeader As IRfcStructure = oRfcFunction.GetStructure("DOC_HEADER")
            Dim oDocItems As IRfcTable = oRfcFunction.GetTable("DOC_ITEMS")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oDocItems.Clear()
            oRETURN.Clear()
            oDocHeader.SetValue("CO_AREA", pKokrs)
            oDocHeader.SetValue("DOCDATE", pBldat)
            oDocHeader.SetValue("POSTGDATE", pBuDat)
            If destination.User Is Nothing Then
                oDocHeader.SetValue("USERNAME", destination.SystemAttributes.User)
            Else
                oDocHeader.SetValue("USERNAME", destination.User)
            End If
            oRfcFunction.SetValue("IGNORE_WARNINGS", "X")
            Dim lRow As SapAcctngStatKeyFiguresDocItem
            Dim lField As SAPCommon.TField
            For Each lRow In pData
                oDocItems.Append()
                For Each lField In lRow.item.Values
                    If lField.FType = "F" Then
                        oDocItems.SetValue(lField.Name, Decimal.Round(CDec(lField.Value), 3))
                    Else
                        oDocItems.SetValue(lField.Name, lField.Value)
                    End If
                Next
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean
            aErr = False
            For i As Integer = 0 To oRETURN.Count - 1
                post = post & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit()
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAcctngRepstPrimCosts")
            post = "Error: Exception in post"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class
