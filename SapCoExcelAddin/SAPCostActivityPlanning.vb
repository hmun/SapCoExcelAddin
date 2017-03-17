' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPCostActivityPlanning
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        Try
            sapcon = aSapCon
            destination = aSapCon.getDestination()
            sapcon.checkCon()
        Catch ex As System.Exception
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
        End Try
    End Sub

    Public Function ReadActivityOutputTot(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection, pContrl As Collection) As String
        ReadActivityOutputTot = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_READACTOUTPUT")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oTotValue As IRfcTable = oRfcFunction.GetTable("TOTVALUE")
            Dim oContrl As IRfcTable = oRfcFunction.GetTable("CONTRL")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)

            Dim lCnt As Integer
            Dim aObjRow As Object
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oTotValue.Append()
                oTotValue.SetValue("VALUE_INDEX", lCnt)
                oContrl.Append()
                oContrl.SetValue("ATTRIB_INDEX", lCnt)
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                ReadActivityOutputTot = "Success"
                For i As Integer = 0 To oTotValue.Count - 1
                    pData.Add(oTotValue(i))
                Next i
                For i As Integer = 0 To oContrl.Count - 1
                    pContrl.Add(oContrl(i))
                Next i
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    ReadActivityOutputTot = ReadActivityOutputTot & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            ReadActivityOutputTot = "Error: Exception in ReadActivityOutputTot"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function ReadPrimCostTot(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection) As String
        ReadPrimCostTot = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_READPRIMCOST")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oTotValue As IRfcTable = oRfcFunction.GetTable("TOTVALUE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)

            Dim lCnt As Integer
            Dim aObjRow As Object
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oTotValue.Append()
                oTotValue.SetValue("VALUE_INDEX", lCnt)
                oTotValue.SetValue("COST_ELEM", lSAPFormat.unpack(aObjRow.Costelem, 10))
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                ReadPrimCostTot = "Success"
                For i As Integer = 0 To oTotValue.Count - 1
                    pData.Add(oTotValue(i))
                Next i
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    ReadPrimCostTot = ReadPrimCostTot & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            ReadPrimCostTot = "Error: Exception in ReadPrimCostTot"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function ReadActivityInputTot(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection) As String
        ReadActivityInputTot = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_READACTINPUT")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oTotValue As IRfcTable = oRfcFunction.GetTable("TOTVALUE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)

            Dim lCnt As Integer
            Dim aObjRow As Object
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oTotValue.Append()
                oTotValue.SetValue("VALUE_INDEX", lCnt)
                oTotValue.SetValue("SEND_CCTR", lSAPFormat.unpack(aObjRow.SCostcenter, 10))
                oTotValue.SetValue("SEND_ACTIVITY", aObjRow.SActtype)
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                ReadActivityInputTot = "Success"
                For i As Integer = 0 To oTotValue.Count - 1
                    pData.Add(oTotValue(i))
                Next i
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    ReadActivityInputTot = ReadActivityInputTot & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            ReadActivityInputTot = "Error: Exception in ReadActivityInputTot"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function PostPrimCostTot(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection) As String
        PostPrimCostTot = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTPRIMCOST")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oTotValue As IRfcTable = oRfcFunction.GetTable("TOTVALUE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)

            Dim lCnt As Integer
            Dim aObjRow As Object
            Dim aDataRow As Collection
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oTotValue.Append()
                oTotValue.SetValue("VALUE_INDEX", lCnt)
                oTotValue.SetValue("COST_ELEM", lSAPFormat.unpack(aObjRow.Costelem, 10))
                '   move the values from the data
                aDataRow = pData(lCnt)
                oTotValue.SetValue("FIX_VALUE", CDbl(aDataRow(1)))
                oTotValue.SetValue("DIST_KEY_FIX_VAL", CStr(aDataRow(2)))
                oTotValue.SetValue("VAR_VALUE", CDbl(aDataRow(3)))
                oTotValue.SetValue("DIST_KEY_VAR_VAL", CStr(aDataRow(4)))
                If CStr(aDataRow(6)) <> "" Then
                    oTotValue.SetValue("FIX_QUAN", aDataRow(5))
                    oTotValue.SetValue("DIST_KEY_FIX_QUAN", CStr(aDataRow(6)))
                End If
                If CStr(aDataRow(8)) <> "" Then
                    oTotValue.SetValue("VAR_QUAN", aDataRow(7))
                    oTotValue.SetValue("DIST_KEY_VAR_QUAN", CStr(aDataRow(8)))
                End If
                If CStr(aDataRow(9)) <> "" Then
                    oTotValue.SetValue("UNIT_OF_MEASURE", CStr(aDataRow(9)))
                End If
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                PostPrimCostTot = "Success"
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit()
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    PostPrimCostTot = PostPrimCostTot & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            PostPrimCostTot = "Error: Exception in PostPrimCostTot"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function PostActivityOutputTot(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection, pContrl As Collection) As String
        PostActivityOutputTot = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTACTOUTPUT")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oTotValue As IRfcTable = oRfcFunction.GetTable("TOTVALUE")
            Dim oContrl As IRfcTable = oRfcFunction.GetTable("CONTRL")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)

            Dim lCnt As Integer
            Dim aObjRow As Object
            Dim aDataRow As Collection
            Dim aCtrlRow As Object
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oTotValue.Append()
                oTotValue.SetValue("VALUE_INDEX", lCnt)
                '   move the values from the data
                aDataRow = pData(lCnt)
                oTotValue.SetValue("UNIT_OF_MEASURE", CStr(aDataRow(1)))
                oTotValue.SetValue("CURRENCY", CStr(aDataRow(2)))
                oTotValue.SetValue("ACTVTY_QTY", CDbl(aDataRow(3)))
                oTotValue.SetValue("DIST_KEY_QUAN", CStr(aDataRow(4)))
                oTotValue.SetValue("ACTVTY_CAPACTY", CDbl(aDataRow(5)))
                oTotValue.SetValue("DIST_KEY_CAPCTY", CStr(aDataRow(6)))
                oTotValue.SetValue("PRICE_FIX", CDbl(aDataRow(7)))
                oTotValue.SetValue("DIST_KEY_PRICE_FIX", CStr(aDataRow(8)))
                oTotValue.SetValue("PRICE_VAR", CDbl(aDataRow(9)))
                oTotValue.SetValue("DIST_KEY_PRICE_VAR", CStr(aDataRow(10)))
                oTotValue.SetValue("PRICE_UNIT", CInt(aDataRow(11)))
                oTotValue.SetValue("EQUIVALENCE_NO", CInt(aDataRow(12)))
                '   move the values from the contrl
                aCtrlRow = pContrl(lCnt)
                oContrl.Append()
                oContrl.SetValue("ATTRIB_INDEX", lCnt)
                oContrl.SetValue("PRICE_INDICATOR", CStr(lSAPFormat.unpack(aCtrlRow(1), 3)))
                oContrl.SetValue("SWITCH_LAYOUT", CStr(aCtrlRow(2)))
                oContrl.SetValue("ALLOC_COST_ELEM", CStr(lSAPFormat.unpack(aCtrlRow(3), 10)))
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                PostActivityOutputTot = "Success"
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit()
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    PostActivityOutputTot = PostActivityOutputTot & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            PostActivityOutputTot = "Error: Exception in PostActivityOutputTot"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function PostActivityInputTot(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection) As String
        PostActivityInputTot = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTACTINPUT")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oTotValue As IRfcTable = oRfcFunction.GetTable("TOTVALUE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)

            Dim lCnt As Integer
            Dim aObjRow As Object
            Dim aDataRow As Collection
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oTotValue.Append()
                oTotValue.SetValue("VALUE_INDEX", lCnt)
                oTotValue.SetValue("SEND_CCTR", lSAPFormat.unpack(aObjRow.SCostcenter, 10))
                oTotValue.SetValue("SEND_ACTIVITY", aObjRow.SActtype)
                '   move the values from the data
                aDataRow = pData(lCnt)
                oTotValue.SetValue("UNIT_OF_MEASURE", CStr(aDataRow(1)))
                oTotValue.SetValue("QUANTITY_FIX", CDbl(aDataRow(2)))
                oTotValue.SetValue("DIST_KEY_QUAN_FIX", CStr(aDataRow(3)))
                oTotValue.SetValue("QUANTITY_VAR", CDbl(aDataRow(4)))
                oTotValue.SetValue("DIST_KEY_QUAN_VAR", CStr(aDataRow(5)))
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                PostActivityInputTot = "Success"
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit()
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    PostActivityInputTot = PostActivityInputTot & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            PostActivityInputTot = "Error: Exception in PostActivityInputTot"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function


    Public Function PostKeyFigure(pCoAre As String, pFiscy As String, pPfrom As String,
                            pPto As String, pVers As String, pCurt As String,
                            pObjects As Collection, pData As Collection) As String
        PostKeyFigure = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTKEYFIGURE")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oPervalue As IRfcTable = oRfcFunction.GetTable("PERVALUE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)

            Dim lCnt As Integer
            Dim aObjRow As Object
            Dim aDataRow As Collection
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oPervalue.Append()
                oPervalue.SetValue("VALUE_INDEX", lCnt)
                '   move the values from the data
                aDataRow = pData(lCnt)
                oPervalue.SetValue("STATKEYFIG", lSAPFormat.unpack(aDataRow(1), 6))
                oPervalue.SetValue("QUANTITY_PER01", CDbl(aDataRow(2)))
                oPervalue.SetValue("QUANTITY_PER02", CDbl(aDataRow(3)))
                oPervalue.SetValue("QUANTITY_PER03", CDbl(aDataRow(4)))
                oPervalue.SetValue("QUANTITY_PER04", CDbl(aDataRow(5)))
                oPervalue.SetValue("QUANTITY_PER05", CDbl(aDataRow(6)))
                oPervalue.SetValue("QUANTITY_PER06", CDbl(aDataRow(7)))
                oPervalue.SetValue("QUANTITY_PER07", CDbl(aDataRow(8)))
                oPervalue.SetValue("QUANTITY_PER08", CDbl(aDataRow(9)))
                oPervalue.SetValue("QUANTITY_PER09", CDbl(aDataRow(10)))
                oPervalue.SetValue("QUANTITY_PER10", CDbl(aDataRow(11)))
                oPervalue.SetValue("QUANTITY_PER11", CDbl(aDataRow(12)))
                oPervalue.SetValue("QUANTITY_PER12", CDbl(aDataRow(13)))
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                PostKeyFigure = "Success"
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit()
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    PostKeyFigure = PostKeyFigure & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            PostKeyFigure = "Error: Exception in PostKeyFigure"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function ReadKeyFigure(pCoAre As String, pFiscy As String, pPfrom As String,
                                    pPto As String, pVers As String, pCurt As String,
                                    pObjects As Collection, pData As Collection) As String
        ReadKeyFigure = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_READKEYFIGURE")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oPervalue As IRfcTable = oRfcFunction.GetTable("PERVALUE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)

            Dim lCnt As Integer
            Dim aObjRow As Object
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oPervalue.Append()
                oPervalue.SetValue("VALUE_INDEX", lCnt)
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                ReadKeyFigure = "Success"
                For i As Integer = 0 To oPervalue.Count - 1
                    pData.Add(oPervalue(i))
                Next i
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    ReadKeyFigure = ReadKeyFigure & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            ReadKeyFigure = "Error: Exception in ReadKeyFigure"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class

