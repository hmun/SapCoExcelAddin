' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPAcctngManCostAlloc
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        Try
            sapcon = aSapCon
            destination = aSapCon.getDestination()
            sapcon.checkCon()
        Catch ex As System.Exception
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAcctngManCostAlloc")
        End Try
    End Sub

    Public Function post(pKokrs As String, pBuDat As Date, pBldat As Date, pData As Collection, pTest As Boolean) As String
        post = ""
        Try
            If pTest Then
                oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_MANUAL_ALLOC_CHECK")
            Else
                oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_MANUAL_ALLOC_POST")
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
            Dim lRow As Object
            For Each lRow In pData
                oDocItems.Append()
                oDocItems.SetValue("SEND_CCTR", lSAPFormat.unpack(lRow.SEND_CCTR, 10))
                oDocItems.SetValue("SEN_ORDER", lSAPFormat.unpack(lRow.SEN_ORDER, 12))
                oDocItems.SetValue("SEN_WBS_EL", CStr(lRow.SEN_WBS_EL))
                oDocItems.SetValue("SEN_NETWRK", lSAPFormat.unpack(lRow.SEN_NETWRK, 12))
                oDocItems.SetValue("SENOPERATN", lSAPFormat.unpack(lRow.SENOPERATN, 4))
                oDocItems.SetValue("SEND_FUNCTION", CStr(lRow.SEND_FUNCTION))
                oDocItems.SetValue("PERSON_NO", lSAPFormat.unpack(lRow.PERSON_NO, 8))
                oDocItems.SetValue("COST_ELEM", lSAPFormat.unpack(lRow.COST_ELEM, 10))
                oDocItems.SetValue("VALUE_TCUR", Decimal.Round(CDec(lRow.VALUE_TCUR), 2))
                oDocItems.SetValue("SEG_TEXT", CStr(lRow.SEG_TEXT))
                oDocItems.SetValue("REC_CCTR", lSAPFormat.unpack(lRow.REC_CCTR, 10))
                oDocItems.SetValue("REC_ORDER", lSAPFormat.unpack(lRow.REC_ORDER, 12))
                oDocItems.SetValue("REC_WBS_EL", CStr(lRow.REC_WBS_EL))
                oDocItems.SetValue("REC_NETWRK", lSAPFormat.unpack(lRow.REC_NETWRK, 12))
                oDocItems.SetValue("RECOPERATN", lSAPFormat.unpack(lRow.RECOPERATN, 4))
                oDocItems.SetValue("REC_FUNCTION", CStr(lRow.REC_FUNCTION))
                oDocItems.SetValue("TRANS_CURR", CStr(lRow.TRANS_CURR))
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
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAcctngManCostAlloc")
            post = "Error: Exception in post"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try

    End Function

End Class
