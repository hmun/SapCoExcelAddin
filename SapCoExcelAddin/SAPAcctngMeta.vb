' Copyright 2025 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPAcctngMeta

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon
    Private aIntPar As SAPCommon.TStr
    Private cName As String = "SAPAcctngMeta"

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

    Public Sub getMeta_AcctngActivityAllocPost(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
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

    Public Sub getMeta_AcctngManCostAllocPost(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {"DOC_HEADER"}
        Dim aImports As String() = {"IGNORE_WARNINGS"}
        Dim aTables As String() = {"DOC_ITEMS", "RETURN", "CUSTOMER_FIELDS"}
        Try
            log.Debug("getMeta_Post - " & "creating Function BAPI_ACC_MANUAL_ALLOC_POST")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_MANUAL_ALLOC_POST")
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

    Public Sub getMeta_AcctngRepstPrimCostsPost(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {"DOC_HEADER"}
        Dim aImports As String() = {"IGNORE_WARNINGS"}
        Dim aTables As String() = {"DOC_ITEMS", "RETURN", "SEND_CRITERIA", "REC_CRITERIA", "CUSTOMER_FIELDS"}
        Try
            log.Debug("getMeta_Post - " & "creating Function BAPI_ACC_PRIMARY_COSTS_POST")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_PRIMARY_COSTS_POST")
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

    Public Sub getMeta_AcctngStatKeyFiguresPost(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {"DOC_HEADER"}
        Dim aImports As String() = {"IGNORE_WARNINGS"}
        Dim aTables As String() = {"DOC_ITEMS", "RETURN", "CUSTOMER_FIELDS"}
        Try
            log.Debug("getMeta_Post - " & "creating Function BAPI_ACC_STAT_KEY_FIG_POST")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_ACC_STAT_KEY_FIG_POST")
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

End Class
