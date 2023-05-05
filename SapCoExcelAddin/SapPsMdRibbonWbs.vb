' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SapPsMdRibbonWbs

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Public Function getGenParameters(ByRef pPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim i As Integer
        log.Debug("SapPsMdRibbonWbs getGenParametrs - " & "reading Parameter")
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP PS Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
            getGenParameters = False
            Exit Function
        End Try
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> "SAPPsMd" And aKey <> "SAPCoMultiple" Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key SAPPsMd or SAPCoMultiple. Check if the current workbook is a valid SAP PS Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
            getGenParameters = False
            Exit Function
        End If
        i = 2
        pPar = New SAPCommon.TStr
        Do While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
            pPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 4).value), pFORMAT:=CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop
        getGenParameters = True
    End Function

    Private Function getIntParameters(ByRef pIntPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim i As Integer

        log.Debug("getIntParameters - " & "reading Parameter")
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter_Int")
        Catch Exc As System.Exception
            MsgBox("No Parameter_Int Sheet in current workbook. Check if the current workbook is a valid SAP PS Md Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
            getIntParameters = False
            Exit Function
        End Try
        i = 2
        pIntPar = New SAPCommon.TStr
        Do
            pIntPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
        ' no obligatory parameters check - we should know what we are doing
        getIntParameters = True
    End Function

    Public Sub exec(ByRef pSapCon As SapCon, Optional pMode As String = "Create")
        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr

        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        Dim aData As Collection

        Dim aRetStr As String
        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        Dim aSAPWBSPI As New SAPWBSPI(pSapCon, aIntPar)

        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Dim aDwsName As String = If(aIntPar.value("WBS_WS", "DATA") <> "", aIntPar.value("WBS_WS", "DATA"), "ProjectDefinition")
        Try
            aDws = aWB.Worksheets(aDwsName)
        Catch Exc As System.Exception
            MsgBox("No " & aDwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Project Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
            Exit Sub
        End Try
        ' Read the Items
        Try
            log.Debug("SapPsMdRibbonWbs.exec - " & "processing data - disabling events, screen update, cursor")
            aDws.Activate()
            Dim aItems As New TData(aIntPar)
            Dim aItem As New TDataRec(aIntPar)
            Dim aKey As String
            Dim j As UInt64
            Dim jMax As UInt64 = 0
            Dim aObjNr As UInt64 = 0
            Dim aLOff As Integer = If(aIntPar.value("WBS_LOFF", "DATA") <> "", CInt(aIntPar.value("WBS_LOFF", "DATA")), 5)
            Dim aLOffCtrl As Integer = If(aIntPar.value("WBS_LOFFCTRL", "DATA") <> "", CInt(aIntPar.value("WBS_LOFFCTRL", "DATA")), 4)
            Dim aDumpObjNr As UInt64 = If(aIntPar.value("WBS_DBG", "DUMPOBJNR") <> "", CInt(aIntPar.value("WBS_DBG", "DUMPOBJNR")), 0)
            Dim aMsgClmn As String = If(aIntPar.value("WBS_COL", "DATAMSG") <> "", aIntPar.value("WBS_COL", "DATAMSG"), "INT-MSG")
            Dim aMsgClmnNr As Integer = 0
            Dim aProjectClmn As String = If(aIntPar.value("WBS_COL", "PROJECT") <> "", aIntPar.value("WBS_COL", "PROJECT"), "I_PROJECT_DEFINITION")
            Dim aProjectClmnNr As Integer = 0
            Dim aOKMsg As String = If(aIntPar.value("WBS_RET", "OKMSG") <> "", aIntPar.value("WBS_RET", "OKMSG"), "OK")
            Dim aSingleMode As String = If(aIntPar.value("WBS", "SINGLE_MODE") <> "", aIntPar.value("WBS", "SINGLE_MODE"), "")
            If pMode = "CreateSingle" Then
                aSingleMode = "X"
            End If
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapCoExcelAddin.Application.EnableEvents = False
            Globals.SapCoExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aLOff + 1
            ' determine the last column and create the fieldlist
            Do
                jMax += 1
                If CStr(aDws.Cells(1, jMax).value) = aMsgClmn Then
                    aMsgClmnNr = jMax
                ElseIf CStr(aDws.Cells(1, jMax).value) = aProjectClmn Then
                    aProjectClmnNr = jMax
                End If
            Loop While CStr(aDws.Cells(aLOff - 4, jMax + 1).value) <> ""
            Dim aProject As String = ""
            aData = New Collection
            j = 1
            Do
                aObjNr += 1
                ' WBS are handled line by line and not in packages. Using aItems is for standardization reasons - Will only contain one item.
                If Left(CStr(aDws.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    For j = 1 To jMax
                        If CStr(aDws.Cells(1, j).value) <> "N/A" And CStr(aDws.Cells(1, j).value) <> "" And CStr(aDws.Cells(1, j).value) <> aMsgClmn And
                            CStr(aDws.Cells(aLOffCtrl + 1, j).value) <> "" And CStr(aDws.Cells(aLOffCtrl + 1, j).value) <> "N" Then
                            aProject = CStr(aDws.Cells(i, aProjectClmnNr).value)
                            aItems.addValue(aKey, CStr(aDws.Cells(aLOff - 4, j).value), CStr(aDws.Cells(i, j).value),
                                    CStr(aDws.Cells(aLOff - 3, j).value), CStr(aDws.Cells(aLOff - 2, j).value), pEmty:=False,
                                    pEmptyChar:="")
                        End If
                    Next
                    Dim aNextProject As String = nextProject(aDws, i, aMsgClmnNr, aProjectClmnNr, aOKMsg)
                    If aProject <> aNextProject Or aSingleMode = "X" Then
                        Dim aTSAP_WbsData
                        If pMode = "Change" Then
                            aTSAP_WbsData = New TSAP_WbsChgData(aPar, aIntPar)
                        Else
                            aTSAP_WbsData = New TSAP_WbsData(aPar, aIntPar)
                        End If
                        If aTSAP_WbsData.fillHeader(aItems) And aTSAP_WbsData.fillData(aItems) Then
                            ' check if we should dump this document
                            If aObjNr = aDumpObjNr Then
                                log.Debug("SapPsMdRibbonWbs.exec - " & "dumping Object Nr " & CStr(aObjNr))
                                aTSAP_WbsData.dumpHeader()
                                aTSAP_WbsData.dumpData()
                            End If
                            ' post the object here
                            If pMode = "Create" Or pMode = "CreateSingle" Then
                                log.Debug("SapPsMdRibbonWbs.exec - " & "calling aSAPWBSPI.createMultiple")
                                aRetStr = aSAPWBSPI.createMultiple(aTSAP_WbsData)
                                log.Debug("SapPsMdRibbonWbs.exec - " & "aSAPWBSPI.createMultiple returned, aRetStr=" & aRetStr)
                                ' message has to be written in all lines that where processed in items
                                For Each aKey In aItems.aTDataDic.Keys
                                    aDws.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                                Next
                            ElseIf pMode = "Change" Then
                                log.Debug("SapPsMdRibbonWbs.exec - " & "calling aSAPWBSPI.changeMultiple")
                                aRetStr = aSAPWBSPI.changeMultiple(aTSAP_WbsData)
                                log.Debug("SapPsMdRibbonWbs.exec - " & "aSAPWBSPI.changeMultiple returned, aRetStr=" & aRetStr)
                                aDws.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                                ' message has to be written in all lines that where processed in items
                                For Each aKey In aItems.aTDataDic.Keys
                                    aDws.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                                Next
                            End If
                        Else
                            log.Warn("SapPsMdRibbonWbs.exec - " & "filling Header or Data in aTSAP_WbsData failed!")
                            aDws.Cells(i, aMsgClmnNr) = "Filling Header or Data in aTSAP_WbsData failed!"
                        End If
                        aItems = New TData(aIntPar)
                    End If
                Else
                    aDws.Cells(i, aMsgClmnNr + 1).Value = "ignored - already processed"
                End If
                i += 1
            Loop While CStr(aDws.Cells(i, 1).Value) <> ""

            log.Debug("SapPsMdRibbonWbs.exec - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapPsMdRibbonWbs.exec failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap AddIn")
            log.Error("SapPsMdRibbonWbs.exec - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Public Sub exec_settle(ByRef pSapCon As SapCon, Optional pMode As String = "Create")
        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr

        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        Dim aData As Collection

        Dim aRetStr As String
        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        Dim aSAPWBSPI As New SAPWBSPI(pSapCon, aIntPar)

        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Dim aDwsName As String = If(aIntPar.value("WBS_WS", "DATA") <> "", aIntPar.value("WBS_WS", "DATA"), "ProjectDefinition")
        Try
            aDws = aWB.Worksheets(aDwsName)
        Catch Exc As System.Exception
            MsgBox("No " & aDwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Project Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
            Exit Sub
        End Try
        ' Read the Items
        Try
            log.Debug("SapPsMdRibbonWbs.exec_settle - " & "processing data - disabling events, screen update, cursor")
            aDws.Activate()
            Dim aItems As New TData(aIntPar)
            Dim aItem As New TDataRec(aIntPar)
            Dim aKey As String
            Dim j As UInt64
            Dim jMax As UInt64 = 0
            Dim aObjNr As UInt64 = 0
            Dim aLOff As Integer = If(aIntPar.value("WBS_LOFF", "DATA") <> "", CInt(aIntPar.value("WBS_LOFF", "DATA")), 5)
            Dim aLOffCtrl As Integer = If(aIntPar.value("WBS_LOFFCTRL", "DATA") <> "", CInt(aIntPar.value("WBS_LOFFCTRL", "DATA")), 4)
            Dim aDumpObjNr As UInt64 = If(aIntPar.value("WBS_DBG", "DUMPOBJNR") <> "", CInt(aIntPar.value("WBS_DBG", "DUMPOBJNR")), 0)
            Dim aMsgClmn As String = If(aIntPar.value("WBS_COL", "SETTLEMSG") <> "", aIntPar.value("WBS_COL", "SETTLEMSG"), "INT-SETTLEMSG")
            Dim aMsgClmnNr As Integer = 0
            Dim aOKMsg As String = If(aIntPar.value("WBS_RET", "OKMSG") <> "", aIntPar.value("WBS_RET", "OKMSG"), "OK")
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapCoExcelAddin.Application.EnableEvents = False
            Globals.SapCoExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aLOff + 1
            ' determine the last column and create the fieldlist
            Do
                jMax += 1
                If CStr(aDws.Cells(1, jMax).value) = aMsgClmn Then
                    aMsgClmnNr = jMax
                End If
            Loop While CStr(aDws.Cells(aLOff - 4, jMax + 1).value) <> ""
            Dim aProject As String = ""
            aData = New Collection
            j = 1
            Do
                aObjNr += 1
                ' WBS are handled line by line and not in packages. Using aItems is for standardization reasons - Will only contain one item.
                If Left(CStr(aDws.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    For j = 1 To jMax
                        If CStr(aDws.Cells(1, j).value) <> "N/A" And CStr(aDws.Cells(1, j).value) <> "" And CStr(aDws.Cells(1, j).value) <> aMsgClmn And
                            CStr(aDws.Cells(aLOffCtrl + 1, j).value) <> "" And CStr(aDws.Cells(aLOffCtrl + 1, j).value) <> "N" Then
                            aItems.addValue(aKey, CStr(aDws.Cells(aLOff - 4, j).value), CStr(aDws.Cells(i, j).value),
                                    CStr(aDws.Cells(aLOff - 3, j).value), CStr(aDws.Cells(aLOff - 2, j).value), pEmty:=False,
                                    pEmptyChar:="")
                        End If
                    Next
                    Dim TSAP_WbsSettleData As New TSAP_WbsSettleData(aPar, aIntPar)
                    If TSAP_WbsSettleData.fillHeader(aItems) Then
                        ' check if we should dump this document
                        If aObjNr = aDumpObjNr Then
                            log.Debug("SapPsMdRibbonWbs.exec_settle - " & "dumping Object Nr " & CStr(aObjNr))
                            TSAP_WbsSettleData.dumpHeader()
                        End If
                        ' post the object here
                        If pMode = "Create" Then
                            log.Debug("SapPsMdRibbonWbs.exec_settle - " & "calling aSAPWBSPI.createMultiple")
                            aRetStr = aSAPWBSPI.createSettlementRule(TSAP_WbsSettleData)
                            log.Debug("SapPsMdRibbonWbs.exec_settle - " & "aSAPWBSPI.createMultiple returned, aRetStr=" & aRetStr)
                            ' message has to be written in all lines that where processed in items
                            For Each aKey In aItems.aTDataDic.Keys
                                aDws.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                            Next
                        ElseIf pMode = "Change" Then
                            ' log.Debug("SapPsMdRibbonWbs.exec_settle - " & "calling aSAPWBSPI.changeSettlementRule")
                            ' aRetStr = aSAPWBSPI.changeSettlementRule(TSAP_WbsSettleData)
                            ' log.Debug("SapPsMdRibbonWbs.exec_settle - " & "aSAPWBSPI.changeSettlementRule returned, aRetStr=" & aRetStr)
                            ' aDws.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                        End If
                    Else
                        log.Warn("SapPsMdRibbonWbs.exec_settle - " & "filling Header or Data in TSAP_WbsSettleData failed!")
                        aDws.Cells(i, aMsgClmnNr) = "Filling Header or Data in TSAP_WbsSettleData failed!"
                    End If
                    aItems = New TData(aIntPar)
                Else
                    aDws.Cells(i, aMsgClmnNr + 1).Value = "ignored - already processed"
                End If
                i += 1
            Loop While CStr(aDws.Cells(i, 1).Value) <> ""

            log.Debug("SapPsMdRibbonWbs.exec_settle - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapPsMdRibbonWbs.exec_settle failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap AddIn")
            log.Error("SapPsMdRibbonWbs.exec_settle - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Public Sub Status(ByRef pSapCon As SapCon, Optional pMode As String = "Set")
        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr

        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        Dim aSAPWBSPI As New SAPWBSPI(pSapCon, aIntPar)

        Dim jMax As UInt64 = 0
        Dim aObjNr As UInt64 = 0
        Dim aWbsLOff As Integer = If(aIntPar.value("WBS_LOFF", "DATA") <> "", CInt(aIntPar.value("WBS_LOFF", "DATA")), 4)
        Dim aLOffCtrl As Integer = If(aIntPar.value("WBS_LOFFCTRL", "DATA") <> "", CInt(aIntPar.value("WBS_LOFFCTRL", "DATA")), 4)
        Dim aDumpObjNr As UInt64 = If(aIntPar.value("WBS_DBG", "DUMPOBJNR") <> "", CInt(aIntPar.value("WBS_DBG", "DUMPOBJNR")), 0)
        Dim aWbsWsName As String = If(aIntPar.value("WBS_WS", "DATA") <> "", aIntPar.value("WBS_WS", "DATA"), "WBS")
        Dim aWbsWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value("WBS_COL", "STATUSMSG") <> "", aIntPar.value("WBS_COL", "STATUSMSG"), "INT-STATUSMSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value("WBS_RET", "OKMSG") <> "", aIntPar.value("WBS_RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aWbsWs = aWB.Worksheets(aWbsWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aWbsWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Project Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap PS Md")
            Exit Sub
        End Try
        parseHeaderLine(aWbsWs, jMax, aMsgClmn, aMsgClmnNr, pHdrLine:=1)
        Try
            log.Debug("SapPsMdRibbonWbs.SetStatus - " & "processing data - disabling events, screen update, cursor")
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapCoExcelAddin.Application.EnableEvents = False
            Globals.SapCoExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aWbsLOff + 1
            Dim aKey As String
            Do
                aObjNr += 1
                If Left(CStr(aWbsWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    Dim aWbsItems As New TData(aIntPar)
                    Dim aTSAP_WbsStatusData As TSAP_WbsGenData
                    If pMode = "Get" Then
                        aTSAP_WbsStatusData = New TSAP_WbsGenData(aPar, aIntPar, aSAPWBSPI, "GetStatus")
                    Else
                        aTSAP_WbsStatusData = New TSAP_WbsGenData(aPar, aIntPar, aSAPWBSPI, "SetStatus")
                    End If
                    ' read DATA
                    aWbsItems.ws_parse_line_simple(aWbsWsName, aWbsLOff, i, jMax, pHdrLine:=1, pUplLine:=aLOffCtrl + 1)
                    If aTSAP_WbsStatusData.fillData(aWbsItems) Then
                        ' check if we should dump this document
                        If aObjNr = aDumpObjNr Then
                            log.Debug("SapPsMdRibbonWbs.exec - " & "dumping Object Nr " & CStr(aObjNr))
                            aTSAP_WbsStatusData.dumpHeader()
                        End If
                        If pMode = "Get" Then
                            log.Debug("SapPsMdRibbonWbs.GetStatus - " & "calling aSAPWBSPI.GetStatus")
                            aRetStr = aSAPWBSPI.GetStatus(aTSAP_WbsStatusData, pOKMsg:=aOKMsg)
                            log.Debug("SapPsMdRibbonWbs.GetStatus - " & "aSAPWBSPI.GetStatus returned, aRetStr=" & aRetStr)
                            aWbsWs.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                            ' output the data now
                            Dim aTData As TData
                            If aTSAP_WbsStatusData.aDataDic.aTDataDic.ContainsKey("E_SYSTEM_STATUS") Then
                                aTData = aTSAP_WbsStatusData.aDataDic.aTDataDic("E_SYSTEM_STATUS")
                                aTData.ws_output_line(aWbsWsName, "", i, jMax, pCoff:=0, pClear:=False, pKey:="")
                            End If
                            If aTSAP_WbsStatusData.aDataDic.aTDataDic.ContainsKey("E_USER_STATUS") Then
                                aTData = aTSAP_WbsStatusData.aDataDic.aTDataDic("E_USER_STATUS")
                                aTData.ws_output_line(aWbsWsName, "", i, jMax, pCoff:=0, pClear:=False, pKey:="")
                            End If
                        ElseIf pMode = "Set" Then
                            log.Debug("SapPsMdRibbonWbs.SetStatus - " & "calling aSAPWBSPI.SetStatus")
                            aRetStr = aSAPWBSPI.SetStatus(aTSAP_WbsStatusData, pOKMsg:=aOKMsg)
                            log.Debug("SapPsMdRibbonWbs.SetStatus - " & "aSAPWBSPI.SetStatus returned, aRetStr=" & aRetStr)
                            aWbsWs.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                        End If
                    End If
                End If
                i += 1
            Loop While Not String.IsNullOrEmpty(CStr(aWbsWs.Cells(i, 1).value))
            log.Debug("SapPsMdRibbonWbs.SetStatus - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapPsMdRibbonWbs.SetStatus failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap AddIn")
            log.Error("SapPsMdRibbonWbs.SetStatus - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Public Sub GetData(ByRef pSapCon As SapCon)
        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr

        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If
        Dim aSAPWBSPI As New SAPWBSPI(pSapCon, pIntPar:=aIntPar)

        Dim jMax As UInt64 = 0
        Dim aWLiLOff As Integer = If(aIntPar.value("LOFF", "WBS_LIST") <> "", CInt(aIntPar.value("LOFF", "WBS_LIST")), 4)
        Dim aWbsLOff As Integer = If(aIntPar.value("LOFF", "WBS_DATA") <> "", CInt(aIntPar.value("LOFF", "WBS_DATA")), 4)
        Dim aWLiWsName As String = If(aIntPar.value("WS", "WBS_LIST") <> "", aIntPar.value("WS", "WBS_LIST"), "WBS_List")
        Dim aWbsWsName As String = If(aIntPar.value("WS", "WBS_DATA") <> "", aIntPar.value("WS", "WBS_DATA"), "WBS_Data")
        Dim aWbsWs As Excel.Worksheet
        Dim aWLiWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value("COL", "WBS_LISTMSG") <> "", aIntPar.value("COL", "WBS_LISTMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aWbsClmnNr As Integer = If(aIntPar.value("COLNR", "WBS_DATA") <> "", CInt(aIntPar.value("COLNR", "WBS_DATA")), 1)
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value("RET", "OKMSG") <> "", aIntPar.value("RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aWLiWs = aWB.Worksheets(aWLiWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aWLiWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Project Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP PS Md")
            Exit Sub
        End Try
        Try
            aWbsWs = aWB.Worksheets(aWbsWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aWbsWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Project Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP PS Md")
            Exit Sub
        End Try
        parseHeaderLine(aWLiWs, jMax, aMsgClmn, aMsgClmnNr, pHdrLine:=aWbsLOff - 3)
        Try
            log.Debug("SapPsMdRibbonWbs.GetData - " & "processing data - disabling events, screen update, cursor")
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapCoExcelAddin.Application.EnableEvents = False
            Globals.SapCoExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aWbsLOff + 1
            Dim aKey As String
            Dim aFirst As Boolean = True
            Dim aClear As Boolean = False
            Do
                If Left(CStr(aWLiWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    aKey = CStr(i)
                    Dim aWLiItems As New TData(aIntPar)
                    aWLiItems.ws_parse_line_simple(aWLiWsName, aWbsLOff, i, jMax)
                    Dim aTSAP_WbsGenData As New TSAP_WbsGenData(aPar, aIntPar, aSAPWBSPI, "GetData")
                    If aTSAP_WbsGenData.fillHeader(aWLiItems) And aTSAP_WbsGenData.fillData(aWLiItems) Then
                        log.Debug("SapPsMdRibbonWbs.GetData - " & "calling aSAPWBSPI.GetData")
                        aRetStr = aSAPWBSPI.GetData(aTSAP_WbsGenData, pOKMsg:=aOKMsg)
                        log.Debug("SapPsMdRibbonWbs.GetData - " & "aSAPWBSPI.GetData returned, aRetStr=" & aRetStr)
                        aWLiWs.Cells(i, aMsgClmnNr) = CStr(aRetStr)
                        If Left(aRetStr, Len(aOKMsg)) = aOKMsg Then
                            If aFirst Then
                                aClear = True
                                aFirst = False
                            Else
                                aClear = False
                            End If
                            Dim aTData As TData
                            If aTSAP_WbsGenData.aDataDic.aTDataDic.ContainsKey("ET_WBS_ELEMENT") Then
                                aTData = aTSAP_WbsGenData.aDataDic.aTDataDic("ET_WBS_ELEMENT")
                                aTData.ws_output(pWs:=aWbsWs, pLoff:=aWbsLOff, pCoff:=0, pPar:=aIntPar, pClear:=aClear, pKey:="")
                            End If
                        End If
                    End If
                End If
                i += 1
            Loop While CStr(aWLiWs.Cells(i, 1).value) <> ""
            log.Debug("SapPsMdRibbonWbs.GetData - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapCoExcelAddin.Application.EnableEvents = True
            Globals.SapCoExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapPsMdRibbonWbs.GetData failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP PS Md AddIn")
            log.Error("SapPsMdRibbonWbs.GetData - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Private Sub parseHeaderLine(ByRef pWs As Excel.Worksheet, ByRef pMaxJ As Integer, Optional pMsgClmn As String = "", Optional ByRef pMsgClmnNr As Integer = 0, Optional pHdrLine As Integer = 1)
        pMaxJ = 0
        Do
            pMaxJ += 1
            If Not String.IsNullOrEmpty(pMsgClmn) And CStr(pWs.Cells(pHdrLine, pMaxJ).value) = pMsgClmn Then
                pMsgClmnNr = pMaxJ
            End If
        Loop While CStr(pWs.Cells(pHdrLine, pMaxJ + 1).value) <> ""
    End Sub

    Function nextProject(ByRef pWs As Excel.Worksheet, pStart As ULong, pMsgClmnNr As Integer, aProjectClmnNr As Integer, pOKMsg As String) As String
        Dim i As ULong = pStart + 1
        nextProject = ""
        Do
            If Left(CStr(pWs.Cells(i, pMsgClmnNr).Value), Len(pOKMsg)) <> pOKMsg Then
                nextProject = CStr(pWs.Cells(i, aProjectClmnNr).Value)
                Exit Function
            End If
            i += 1
        Loop While CStr(pWs.Cells(i, 1).Value) <> ""
    End Function

End Class
