Imports SAP.Middleware.Connector

Public Class SAPCOPAActuals
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        Try
            sapcon = aSapCon
            destination = aSapCon.getDestination()
            sapcon.checkCon()
        Catch ex As System.Exception
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCOPAActuals")
        End Try
    End Sub

    Public Function PostCostingBasedData(pOperatingConcern As String, pData As Collection, Optional pCheck As Boolean = False) As String
        PostCostingBasedData = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COPAACTUALS_POSTCOSTDATA")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            Dim oInputData As IRfcTable = oRfcFunction.GetTable("INPUTDATA")
            Dim oFieldList As IRfcTable = oRfcFunction.GetTable("FIELDLIST")
            oInputData.Clear()
            oFieldList.Clear()
            oRETURN.Clear()
            If pCheck Then
                oRfcFunction.SetValue("TESTRUN", "X")
            Else
                oRfcFunction.SetValue("TESTRUN", "")
            End If
            oRfcFunction.SetValue("OPERATINGCONCERN", pOperatingConcern)

            Dim aRow As Object
            Dim aItem As Object
            Dim lCnt As Integer = 0
            For Each aRow In pData
                lCnt = lCnt + 1
                For Each aItem In aRow
                    oInputData.Append()
                    oInputData.SetValue("RECORD_ID", lCnt)
                    oInputData.SetValue("FIELDNAME", aItem.gFIELDNAME)
                    If aItem.gCURRENCY IsNot Nothing And aItem.gCURRENCY <> "" Then
                        oInputData.SetValue("CURRENCY", aItem.gCURRENCY)
                        oInputData.SetValue("VALUE", Decimal.Round(CDec(aItem.gVALUE), 2))
                    Else
                        oInputData.SetValue("VALUE", aItem.gVALUE)
                    End If
                    If lCnt = 1 Then
                        oFieldList.Append()
                        oFieldList.SetValue("FIELDNAME", aItem.gFIELDNAME)
                    End If
                Next aItem
            Next aRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean
            aErr = False
            For i As Integer = 0 To oRETURN.Count - 1
                PostCostingBasedData = PostCostingBasedData & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                PostCostingBasedData = "Success"
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit()
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCOPAActuals")
            PostCostingBasedData = "Error: Exception in PostCostingBasedData"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class
