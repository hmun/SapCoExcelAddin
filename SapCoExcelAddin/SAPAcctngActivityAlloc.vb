﻿' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPAcctngActivityAlloc
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon
    Private aIntPar As SAPCommon.TStr

    Sub New(aSapCon As SapCon, ByRef pIntPar As SAPCommon.TStr)
        aIntPar = pIntPar
        Try
            sapcon = aSapCon
            aSapCon.getDestination(destination)
            sapcon.checkCon()
        Catch ex As System.Exception
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAcctngActivityAlloc")
        End Try
    End Sub

    Public Function post(pKokrs As String, pBuDat As Date, pBldat As Date, pData As Collection, pTest As Boolean) As String
        post = ""
        Try
            If pTest Then
                oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_ACTIVITY_ALLOC_CHECK")
            Else
                oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_ACTIVITY_ALLOC_POST")
            End If
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat(pIntPar:=aIntPar)
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
                oDocItems.SetValue("PERSON_NO", CInt(lRow.PERSON_NO))
                oDocItems.SetValue("ACTTYPE", CStr(lRow.ACTTYPE))
                oDocItems.SetValue("ACTVTY_QTY", Decimal.Round(CDec(lRow.ACTVTY_QTY), 3))
                oDocItems.SetValue("SEG_TEXT", CStr(lRow.SEG_TEXT))
                oDocItems.SetValue("REC_WBS_EL", CStr(lRow.REC_WBS_EL))
                oDocItems.SetValue("REC_NETWRK", lSAPFormat.unpack(lRow.REC_NETWRK, 12))
                oDocItems.SetValue("RECOPERATN", lSAPFormat.unpack(lRow.RECOPERATN, 4))
                oDocItems.SetValue("REC_ORDER", lSAPFormat.unpack(lRow.REC_ORDER, 12))
                oDocItems.SetValue("REC_CCTR", lSAPFormat.unpack(lRow.REC_CCTR, 10))
                If CStr(lRow.REC_FUNCTION) <> "" Then
                    oDocItems.SetValue("REC_FUNCTION", CStr(lRow.REC_FUNCTION))
                End If
                If CDec(lRow.PRICE) <> 0 Then
                    oDocItems.SetValue("PRICE", Decimal.Round(CDec(lRow.PRICE), 2))
                End If
                If CDec(lRow.PRICE_FIX) <> 0 Then
                    oDocItems.SetValue("PRICE_FIX", Decimal.Round(CDec(lRow.PRICE_FIX), 2))
                End If
                If CDec(lRow.PRICE_VAR) <> 0 Then
                    oDocItems.SetValue("PRICE_VAR", Decimal.Round(CDec(lRow.PRICE_VAR), 2))
                End If
                If CInt(lRow.PRICE_UNIT) <> 0 Then
                    oDocItems.SetValue("PRICE_UNIT", CInt(lRow.PRICE_UNIT))
                End If
                If CStr(lRow.CURR) <> "" Then
                    oDocItems.SetValue("CURRENCY", CStr(lRow.CURR))
                End If
                If lRow.VALUE_TOTAL <> 0 Then
                    oDocItems.SetValue("VALUE_TOTAL", Decimal.Round(CDec(lRow.VALUE_TOTAL), 2))
                End If
                If lRow.VALUE_FIX <> 0 Then
                    oDocItems.SetValue("VALUE_FIX", Decimal.Round(CDec(lRow.VALUE_FIX), 2))
                End If
                If lRow.VALUE_VAR <> 0 Then
                    oDocItems.SetValue("VALUE_VAR", Decimal.Round(CDec(lRow.VALUE_VAR), 2))
                End If
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
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPAcctngActivityAlloc")
            post = "Error: Exception in post"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function
End Class
