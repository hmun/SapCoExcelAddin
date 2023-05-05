' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPWBSPI
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
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWBSPI")
        End Try
    End Sub

    Private Sub addToFieldArray(pArrayName As String, pFieldName As String, ByRef pFieldsDic As Dictionary(Of String, String()))
        Dim aArray As String()
        If pFieldsDic.ContainsKey(pArrayName) Then
            aArray = pFieldsDic(pArrayName)
            Array.Resize(aArray, aArray.Length + 1)
            aArray(aArray.Length - 1) = pFieldName
            pFieldsDic.Remove(pArrayName)
            pFieldsDic.Add(pArrayName, aArray)
        Else
            aArray = {pFieldName}
            pFieldsDic.Add(pArrayName, aArray)
        End If
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

    Public Sub getMeta_SetStatus(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {}
        Dim aImports As String() = {}
        Dim aTables As String() = {"I_WBS_SYSTEM_STATUS", "I_WBS_USER_STATUS"}
        Try
            log.Debug("getMeta_SetStatus - " & "creating Function BAPI_BUS2054_SET_STATUS")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2054_SET_STATUS")
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
            log.Error("getMeta_SetStatus - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWBSPI")
        Finally
            log.Debug("getMeta_GetDetail - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Sub getMeta_GetStatus(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {}
        Dim aImports As String() = {}
        Dim aTables As String() = {"I_WBS_ELEMENTS", "E_SYSTEM_STATUS", "E_USER_STATUS"}
        Try
            log.Debug("getMeta_GetStatus - " & "creating Function BAPI_BUS2054_GET_STATUS")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2054_GET_STATUS")
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
            log.Error("getMeta_GetStatus - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWBSPI")
        Finally
            log.Debug("getMeta_GetDetail - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Sub getMeta_GetData(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {}
        Dim aImports As String() = {"I_PROJECT_DEFINITION", "I_LANGUAGE", "I_MAX_ROWS"}
        Dim aTables As String() = {"IT_WBS_ELEMENT", "ET_WBS_ELEMENT", "ET_RETURN", "EXTENSIONIN", "EXTENSIONOUT"}
        Try
            log.Debug("getMeta_GetData - " & "creating Function BAPI_BUS2054_GETDATA")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2054_GETDATA")
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
            log.Error("getMeta_GetStatus - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWBSPI")
        Finally
            log.Debug("getMeta_GetDetail - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Function createMultiple(pData As TSAP_WbsData, Optional pOKMsg As String = "OK") As String
        createMultiple = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2054_CREATE_MULTI")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            Dim oIT_WBS_ELEMENT As IRfcTable = oRfcFunction.GetTable("IT_WBS_ELEMENT")
            Dim oEXTENSIONIN As IRfcTable = oRfcFunction.GetTable("EXTENSIONIN")
            oRETURN.Clear()
            oIT_WBS_ELEMENT.Clear()
            oEXTENSIONIN.Clear()

            Dim aSAPBapiPS As New SAPBapiPS(sapcon)
            aSAPBapiPS.initialization()

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
            ' set the data values
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            For Each aKvP In pData.aData.aTDataDic
                Dim oIT_WBS_ELEMENTAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "IT_WBS_ELEMENT"
                            If Not oIT_WBS_ELEMENTAppended Then
                                oIT_WBS_ELEMENT.Append()
                                oIT_WBS_ELEMENTAppended = True
                            End If
                            oIT_WBS_ELEMENT.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
            ' Fill Extension fields
            Dim oEXTENSIONINAppended As Boolean = False
            For Each aKvP In pData.aExt.aTDataDic
                aTDataRec = aKvP.Value
                Dim aCustFields As IEnumerable(Of String)
                aCustFields = fillCustomerFields(aTDataRec)
                oEXTENSIONIN.Append()
                oEXTENSIONIN.SetValue("STRUCTURE", "BAPI_TE_WBS_ELEMENT")
                oEXTENSIONIN.SetValue("VALUEPART1", aCustFields.ElementAt(0))
                oEXTENSIONIN.SetValue("VALUEPART2", aCustFields.ElementAt(1))
                oEXTENSIONIN.SetValue("VALUEPART3", aCustFields.ElementAt(2))
                oEXTENSIONIN.SetValue("VALUEPART4", aCustFields.ElementAt(3))
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                createMultiple = createMultiple & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Dim aPreCommitRet As String
                aPreCommitRet = aSAPBapiPS.precommit
                If aPreCommitRet <> "" Then
                    createMultiple = createMultiple & ";" & aPreCommitRet
                    If Left(aPreCommitRet, 6) = "Error:" Then
                        aPreComErr = True
                    End If
                End If
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            createMultiple = If(createMultiple = "", pOKMsg, If(aPreComErr = False And aErr = False, pOKMsg & createMultiple, "Error" & createMultiple))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWBSPI")
            createMultiple = "Error: Exception in createMultiple"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function changeMultiple(pData As TSAP_WbsChgData, Optional pOKMsg As String = "OK") As String
        changeMultiple = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2054_CHANGE_MULTI")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            Dim oIT_WBS_ELEMENT As IRfcTable = oRfcFunction.GetTable("IT_WBS_ELEMENT")
            Dim oIT_UPDATE_WBS_ELEMENT As IRfcTable = oRfcFunction.GetTable("IT_UPDATE_WBS_ELEMENT")
            Dim oEXTENSIONIN As IRfcTable = oRfcFunction.GetTable("EXTENSIONIN")
            oRETURN.Clear()
            oIT_WBS_ELEMENT.Clear()
            oIT_UPDATE_WBS_ELEMENT.Clear()
            oEXTENSIONIN.Clear()

            Dim aSAPBapiPS As New SAPBapiPS(sapcon)
            aSAPBapiPS.initialization()

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
            ' set the data values
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            For Each aKvP In pData.aData.aTDataDic
                Dim oIT_WBS_ELEMENTAppended As Boolean = False
                Dim oIT_UPDATE_WBS_ELEMENTAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "IT_WBS_ELEMENT"
                            If Not oIT_WBS_ELEMENTAppended Then
                                oIT_WBS_ELEMENT.Append()
                                oIT_WBS_ELEMENTAppended = True
                            End If
                            oIT_WBS_ELEMENT.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                        Case "IT_UPDATE_WBS_ELEMENT"
                            If Not oIT_UPDATE_WBS_ELEMENTAppended Then
                                oIT_UPDATE_WBS_ELEMENT.Append()
                                oIT_UPDATE_WBS_ELEMENTAppended = True
                            End If
                            oIT_UPDATE_WBS_ELEMENT.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
            ' Fill Extension fields
            Dim oEXTENSIONINAppended As Boolean = False
            For Each aKvP In pData.aExt.aTDataDic
                aTDataRec = aKvP.Value
                Dim aCustFields As IEnumerable(Of String)
                aCustFields = fillCustomerFields(aTDataRec)
                oEXTENSIONIN.Append()
                oEXTENSIONIN.SetValue("STRUCTURE", "BAPI_TE_WBS_ELEMENT")
                oEXTENSIONIN.SetValue("VALUEPART1", aCustFields.ElementAt(0))
                oEXTENSIONIN.SetValue("VALUEPART2", aCustFields.ElementAt(1))
                oEXTENSIONIN.SetValue("VALUEPART3", aCustFields.ElementAt(2))
                oEXTENSIONIN.SetValue("VALUEPART4", aCustFields.ElementAt(3))
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                changeMultiple = changeMultiple & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Dim aPreCommitRet As String
                aPreCommitRet = aSAPBapiPS.precommit
                If aPreCommitRet <> "" Then
                    changeMultiple = changeMultiple & ";" & aPreCommitRet
                    If Left(aPreCommitRet, 6) = "Error:" Then
                        aPreComErr = True
                    End If
                End If
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            changeMultiple = If(changeMultiple = "", pOKMsg, If(aPreComErr = False And aErr = False, pOKMsg & changeMultiple, "Error" & changeMultiple))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWBSPI")
            changeMultiple = "Error: Exception in changeMultiple"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Function fillCustomerFields(pExtInfo As TDataRec) As IEnumerable(Of String)
        Dim aSAPFormat As New SAPFormat(aIntPar)
        Dim aExtension As New SAPCommon.SapExtension(aIntPar)
        Dim aTStrRec As SAPCommon.TStrRec
        For Each aTStrRec In pExtInfo.aTDataRecCol
            aExtension.addField(aTStrRec)
        Next
        aExtension.addString(aSAPFormat.pspid(pExtInfo.getWbs, 18), 0, 24)
        fillCustomerFields = aExtension.getArray()
    End Function

    Public Function createSettlementRule(pData As TSAP_WbsSettleData, Optional pOKMsg As String = "OK") As String
        createSettlementRule = ""
        Dim aSAPFormat As New SAPFormat(aIntPar)
        Try
            oRfcFunction = destination.Repository.CreateFunction("ZPS_KSRG_WBS")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            oRETURN.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            ' use local Version of the SapFormat.pspid (the common does not support the mask strings)
            If pData.aHdrRec.aTDataRecCol.Count <> 3 Then
                createSettlementRule = pOKMsg & "; not relevant"
                Exit Function
            End If
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    If String.IsNullOrEmpty(aTStrRec.Value) Then
                        createSettlementRule = pOKMsg & "; not relevant"
                        Exit Function
                    Else
                        If Left(aTStrRec.Format, 1) = "P" Then
                            oRfcFunction.SetValue(aTStrRec.Fieldname, aSAPFormat.pspid(aTStrRec.Value, 18))
                        Else
                            oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                        End If
                    End If

                End If
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                createSettlementRule = createSettlementRule & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            createSettlementRule = If(createSettlementRule = "", pOKMsg, If(aErr = False, pOKMsg & createSettlementRule, "Error" & createSettlementRule))
        Catch SapEx As SAP.Middleware.Connector.RfcAbapMessageException
            createSettlementRule = "Error; " & SapEx.AbapMessageType & "-" & SapEx.AbapMessageClass & "-" & SapEx.AbapMessageNumber & ": " & SapEx.Message
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWBSPI")
            createSettlementRule = "Error: Exception in createSettlementRule"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function getSystemStatus(pData As TSAP_WbsChgData, Optional pOKMsg As String = "OK") As String
        getSystemStatus = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2054_GET_STATUS")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            Dim oIT_WBS_ELEMENT As IRfcTable = oRfcFunction.GetTable("IT_WBS_ELEMENT")
            Dim oIT_UPDATE_WBS_ELEMENT As IRfcTable = oRfcFunction.GetTable("IT_UPDATE_WBS_ELEMENT")
            Dim oEXTENSIONIN As IRfcTable = oRfcFunction.GetTable("EXTENSIONIN")
            oRETURN.Clear()
            oIT_WBS_ELEMENT.Clear()
            oIT_UPDATE_WBS_ELEMENT.Clear()
            oEXTENSIONIN.Clear()

            Dim aSAPBapiPS As New SAPBapiPS(sapcon)
            aSAPBapiPS.initialization()

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
            ' set the data values
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            For Each aKvP In pData.aData.aTDataDic
                Dim oIT_WBS_ELEMENTAppended As Boolean = False
                Dim oIT_UPDATE_WBS_ELEMENTAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "IT_WBS_ELEMENT"
                            If Not oIT_WBS_ELEMENTAppended Then
                                oIT_WBS_ELEMENT.Append()
                                oIT_WBS_ELEMENTAppended = True
                            End If
                            oIT_WBS_ELEMENT.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                        Case "IT_UPDATE_WBS_ELEMENT"
                            If Not oIT_UPDATE_WBS_ELEMENTAppended Then
                                oIT_UPDATE_WBS_ELEMENT.Append()
                                oIT_UPDATE_WBS_ELEMENTAppended = True
                            End If
                            oIT_UPDATE_WBS_ELEMENT.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
            ' Fill Extension fields
            Dim oEXTENSIONINAppended As Boolean = False
            For Each aKvP In pData.aExt.aTDataDic
                aTDataRec = aKvP.Value
                Dim aCustFields As IEnumerable(Of String)
                aCustFields = fillCustomerFields(aTDataRec)
                oEXTENSIONIN.Append()
                oEXTENSIONIN.SetValue("STRUCTURE", "BAPI_TE_WBS_ELEMENT")
                oEXTENSIONIN.SetValue("VALUEPART1", aCustFields.ElementAt(0))
                oEXTENSIONIN.SetValue("VALUEPART2", aCustFields.ElementAt(1))
                oEXTENSIONIN.SetValue("VALUEPART3", aCustFields.ElementAt(2))
                oEXTENSIONIN.SetValue("VALUEPART4", aCustFields.ElementAt(3))
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                getSystemStatus = getSystemStatus & ";" & oRETURN(i).GetValue("MESSAGE")
                If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Dim aPreCommitRet As String
                aPreCommitRet = aSAPBapiPS.precommit
                If aPreCommitRet <> "" Then
                    getSystemStatus = getSystemStatus & ";" & aPreCommitRet
                    If Left(aPreCommitRet, 6) = "Error:" Then
                        aPreComErr = True
                    End If
                End If
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            getSystemStatus = If(getSystemStatus = "", pOKMsg, If(aPreComErr = False And aErr = False, pOKMsg & getSystemStatus, "Error" & getSystemStatus))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWBSPI")
            getSystemStatus = "Error: Exception in getSystemStatus"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function SetStatus(pData As TSAP_WbsGenData, Optional pOKMsg As String = "OK") As String
        SetStatus = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2054_SET_STATUS")
            RfcSessionManager.BeginContext(destination)
            Dim oI_WBS_SYSTEM_STATUS As IRfcTable = oRfcFunction.GetTable("I_WBS_SYSTEM_STATUS")
            Dim oI_WBS_USER_STATUS As IRfcTable = oRfcFunction.GetTable("I_WBS_USER_STATUS")
            Dim oE_RESULT As IRfcTable = oRfcFunction.GetTable("E_RESULT")
            oI_WBS_SYSTEM_STATUS.Clear()
            oI_WBS_USER_STATUS.Clear()
            oE_RESULT.Clear()

            Dim aSAPBapiPS As New SAPBapiPS(sapcon)
            aSAPBapiPS.initialization()

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
            ' set the data values
            pData.aDataDic.to_IRfcTable(pKey:="I_WBS_SYSTEM_STATUS", pIRfcTable:=oI_WBS_SYSTEM_STATUS)
            pData.aDataDic.to_IRfcTable(pKey:="I_WBS_USER_STATUS", pIRfcTable:=oI_WBS_USER_STATUS)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim sRETURN As IRfcStructure = oRfcFunction.GetStructure("RETURN")
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            SetStatus = SetStatus & ";" & sRETURN.GetValue("MESSAGE")
            If sRETURN.GetValue("TYPE") = "E" Then
                aErr = True
            End If
            For i As Integer = 0 To oE_RESULT.Count - 1
                SetStatus = SetStatus & ";" & oE_RESULT(i).GetValue("STATUS_ACTION") & "-" & oE_RESULT(i).GetValue("STATUS_TYPE") & "-" & oE_RESULT(i).GetValue("MESSAGE_TEXT")
            Next i
            If aErr = False Then
                Dim aPreCommitRet As String
                aPreCommitRet = aSAPBapiPS.precommit
                If aPreCommitRet <> "" Then
                    SetStatus = SetStatus & ";" & aPreCommitRet
                    If Left(aPreCommitRet, 6) = "Error:" Then
                        aPreComErr = True
                    End If
                End If
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            SetStatus = If(SetStatus = "", pOKMsg, If(aPreComErr = False And aErr = False, pOKMsg & SetStatus, "Error" & SetStatus))

        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWBSPI")
            SetStatus = "Error: Exception in SetStatus"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function GetStatus(pData As TSAP_WbsGenData, Optional pOKMsg As String = "OK") As String
        GetStatus = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2054_GET_STATUS")
            RfcSessionManager.BeginContext(destination)
            Dim oI_WBS_ELEMENTS As IRfcTable = oRfcFunction.GetTable("I_WBS_ELEMENTS")
            Dim oE_SYSTEM_STATUS As IRfcTable = oRfcFunction.GetTable("E_SYSTEM_STATUS")
            Dim oE_USER_STATUS As IRfcTable = oRfcFunction.GetTable("E_USER_STATUS")
            Dim oE_RESULT As IRfcTable = oRfcFunction.GetTable("E_RESULT")
            oI_WBS_ELEMENTS.Clear()
            oE_SYSTEM_STATUS.Clear()
            oE_USER_STATUS.Clear()
            oE_RESULT.Clear()

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
            ' set the data values
            pData.aDataDic.to_IRfcTable(pKey:="I_WBS_ELEMENTS", pIRfcTable:=oI_WBS_ELEMENTS)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim sRETURN As IRfcStructure = oRfcFunction.GetStructure("RETURN")
            Dim aErr As Boolean = False
            Dim aPreComErr As Boolean = False
            GetStatus = GetStatus & ";" & sRETURN.GetValue("MESSAGE")
            If sRETURN.GetValue("TYPE") = "E" Then
                aErr = True
            End If
            If aErr = False Then
                ' return the system status
                pData.aDataDic.addValues(oTable:=oE_SYSTEM_STATUS, pStrucName:="E_SYSTEM_STATUS")
                ' return the user status
                pData.aDataDic.addValues(oTable:=oE_USER_STATUS, pStrucName:="E_USER_STATUS")
            End If
            GetStatus = If(GetStatus = "", pOKMsg, If(aPreComErr = False And aErr = False, pOKMsg & GetStatus, "Error" & GetStatus))

        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWBSPI")
            GetStatus = "Error: Exception in GetStatus"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function GetData(pData As TSAP_WbsGenData, Optional pOKMsg As String = "OK", Optional pCheck As Boolean = False) As String
        GetData = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_BUS2054_GETDATA")
            RfcSessionManager.BeginContext(destination)
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("ET_RETURN")
            Dim oIT_WBS_ELEMENT As IRfcTable = oRfcFunction.GetTable("IT_WBS_ELEMENT")
            Dim oET_WBS_ELEMENT As IRfcTable = oRfcFunction.GetTable("ET_WBS_ELEMENT")
            Dim oEXTENSIONIN As IRfcTable = oRfcFunction.GetTable("EXTENSIONIN")
            Dim oEXTENSIONOUT As IRfcTable = oRfcFunction.GetTable("EXTENSIONOUT")
            oIT_WBS_ELEMENT.Clear()
            oET_WBS_ELEMENT.Clear()
            oEXTENSIONIN.Clear()
            oEXTENSIONOUT.Clear()
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
            ' call the BAPI
            oRfcFunction.Invoke(destination)

            Dim aErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                If oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    GetData = GetData & ";" & oRETURN(i).GetValue("MESSAGE")
                    If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "W" Then
                        aErr = True
                    End If
                End If
            Next i
            GetData = If(GetData = "", pOKMsg, If(aErr = False, pOKMsg & GetData, "Error" & GetData))

            If aErr = False Then
                ' process the return tables
                pData.aDataDic.addValues(oTable:=oIT_WBS_ELEMENT, pStrucName:="IT_WBS_ELEMENT")
                pData.aDataDic.addValues(oTable:=oET_WBS_ELEMENT, pStrucName:="ET_WBS_ELEMENT")
                pData.aDataDic.addValues(oTable:=oEXTENSIONIN, pStrucName:="EXTENSIONIN")
                pData.aDataDic.addValues(oTable:=oEXTENSIONOUT, pStrucName:="EXTENSIONOUT")
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPWBSPI")
            GetData = "Error: Exception in GetData"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class
