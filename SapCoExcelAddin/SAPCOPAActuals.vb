Imports SAP.Middleware.Connector

Public Class SAPCOPAActuals
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon
    Private aIntPar As SAPCommon.TStr
    Private cName As String = "SAPCOPAActuals"

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

    Private Sub addToStrucDic(pArrayName As String, pRfcStructureMetadata As RfcStructureMetadata, ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        If pStrucDic.ContainsKey(pArrayName) Then
            pStrucDic.Remove(pArrayName)
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Private Sub addToFieldDic(pArrayName As String, pRfcStructureMetadata As RfcParameterMetadata, ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata))
        If pFieldDic.ContainsKey(pArrayName) Then
            pFieldDic.Remove(pArrayName)
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Public Sub getMeta_Post(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {"DOC_HEADER"}
        Dim aImports As String() = {"IGNORE_WARNINGS"}
        Dim aTables As String() = {"DOC_ITEMS", "RETURN", "CRITERIA", "CUSTOMER_FIELDS"}
        Try
            log.Debug("getMeta_Post - " & "creating Function BAPI_ACC_ACTIVITY_ALLOC_POST")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_ACTIVITY_ALLOC_POST")
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To aImports.Length - 1
                addToFieldDic("I|" & aImports(s), oRfcFunction.Metadata.Item(aImports(s)), pFieldDic)
            Next
            ' Import Strcutures
            For s As Integer = 0 To aStructures.Length - 1
                oStructure = oRfcFunction.GetStructure(aStructures(s))
                addToStrucDic("S|" & aStructures(s), oStructure.Metadata, pStrucDic)
            Next
            For s As Integer = 0 To aTables.Length - 1
                oTable = oRfcFunction.GetTable(aTables(s))
                addToStrucDic("T|" & aTables(s), oTable.Metadata.LineType, pStrucDic)
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_Post - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, cName)
        Finally
            log.Debug("getMeta_Post - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Function Post(pData As TSAP_Data_CoPa, Optional pOKMsg As String = "OK", Optional pCheck As Boolean = False) As String
        Post = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COPAACTUALS_POSTCOSTDATA")
            If pCheck Then
                oRfcFunction.SetValue("TESTRUN", "X")
            Else
                oRfcFunction.SetValue("TESTRUN", "")
            End If
            RfcSessionManager.BeginContext(destination)
            Dim oINPUTDATA As IRfcTable = oRfcFunction.GetTable("INPUTDATA")
            Dim oFIELDLIST As IRfcTable = oRfcFunction.GetTable("FIELDLIST")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oINPUTDATA.Clear()
            oFIELDLIST.Clear()
            oRETURN.Clear()

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
            ' set the table fields
            pData.aDataDic.to_IRfcTable(pKey:="INPUTDATA", pIRfcTable:=oINPUTDATA)
            pData.aDataDic.to_IRfcTable(pKey:="FIELDLIST", pIRfcTable:=oFIELDLIST)
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

    Public Function PostCostingBasedData(pOperatingConcern As String, pData As Collection, Optional pCheck As Boolean = False) As String
        PostCostingBasedData = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COPAACTUALS_POSTCOSTDATA")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat(pIntPar:=aIntPar)
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
                        oInputData.SetValue("VALUE", CStr(Decimal.Round(CDec(aItem.gVALUE), 2)))
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
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, cName)
            PostCostingBasedData = "Error: Exception in PostCostingBasedData"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class
