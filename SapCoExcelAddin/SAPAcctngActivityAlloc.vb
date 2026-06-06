' Copyright 2025 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPAcctngActivityAlloc
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon
    Private aIntPar As SAPCommon.TStr
    Private cName As String = "SAPAcctngActivityAlloc"

    Sub New(aSapCon As SapCon, ByRef pIntPar As SAPCommon.TStr)
        Try
            log.Debug("New - " & "checking connection")
            sapcon = aSapCon
            aIntPar = pIntPar
            aSapCon.getDestination(destination)
            sapcon.checkCon()
        Catch ex As System.Exception
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, cName)
        End Try
    End Sub

    Public Function Post(pData As TSAP_Data_Co, Optional pOKMsg As String = "OK", Optional pCheck As Boolean = False) As String
        Post = ""
        Try
            If pCheck Then
                oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_ACTIVITY_ALLOC_CHECK")
            Else
                oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_ACTIVITY_ALLOC_POST")
            End If
            RfcSessionManager.BeginContext(destination)
            Dim oDOC_ITEMS As IRfcTable = oRfcFunction.GetTable("DOC_ITEMS")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            Dim oCRITERIA As IRfcTable = oRfcFunction.GetTable("CRITERIA")
            Dim oCUSTOMER_FIELDS As IRfcTable = oRfcFunction.GetTable("CUSTOMER_FIELDS")
            oDOC_ITEMS.Clear()
            oRETURN.Clear()
            oCRITERIA.Clear()
            oCUSTOMER_FIELDS.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            oStruc = oRfcFunction.GetStructure("DOC_HEADER")
            If destination.User Is Nothing Then
                oStruc.SetValue("USERNAME", destination.SystemAttributes.User)
            Else
                oStruc.SetValue("USERNAME", destination.User)
            End If
            ' set the table fields
            pData.aDataDic.to_IRfcTable(pKey:="DOC_ITEMS", pIRfcTable:=oDOC_ITEMS)
            pData.aDataDic.to_IRfcTable(pKey:="CRITERIA", pIRfcTable:=oCRITERIA)
            pData.aDataDic.to_IRfcTable(pKey:="CUSTOMER_FIELDS", pIRfcTable:=oCUSTOMER_FIELDS)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                If oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    Post = Post & ";" & oRETURN(i).GetValue("MESSAGE")
                    If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "W" Then
                        aErr = True
                    End If
                End If
            Next i
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            Post = If(Post = "", pOKMsg, If(aErr = False, pOKMsg & Post, "Error" & Post))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, cName)
            Post = "Error: Exception in Post"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class
