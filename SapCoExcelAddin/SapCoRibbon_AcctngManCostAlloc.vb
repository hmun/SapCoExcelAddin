Public Class SapCoRibbon_AcctngManCostAlloc
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Public Function getGenParameters(ByRef pPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim aName As String
        Dim i As Integer
        log.Debug("SapCoRibbon_AcctngManCostAlloc getGenParametrs - " & "reading Parameter")
        aWB = Globals.ThisAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SapCo Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCo")
            getGenParameters = False
            Exit Function
        End Try
        aName = "SAPAcctngManCostAlloc"
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> aName And aKey <> "SAPCoMultiple" Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key " & aName & ". Check if the current workbook is a valid SapCo Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCo")
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
        aWB = Globals.ThisAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter_Int")
        Catch Exc As System.Exception
            MsgBox("No Parameter_Int Sheet in current workbook. Check if the current workbook is a valid SapCo Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCo")
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

    Public Sub Post(ByRef pSapCon As SapCon, Optional pCheck As Boolean = False)
        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr
        Dim aType As String = "CA"

        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If
        Dim aSAPAcctngManCostAlloc As New SAPAcctngManCostAlloc(pSapCon, aIntPar)
        Dim aSAPAcctngMeta As New SAPAcctngMeta(pSapCon, aIntPar)

        Dim jMax As UInt64 = 0
        Dim aCoLOff As Integer = If(aIntPar.value(aType & "_" & "LOFF", "DATA") <> "", CInt(aIntPar.value(aType & "_" & "LOFF", "DATA")), 4)
        Dim aHdrLOff As Integer = If(aIntPar.value(aType & "_" & "LOFF", "HEAD") <> "", CInt(aIntPar.value(aType & "_" & "LOFF", "HEAD")), aCoLOff - 4)
        Dim aCoWsName As String = If(aIntPar.value(aType & "_" & "WS", "DATA") <> "", aIntPar.value(aType & "_" & "WS", "DATA"), "Data")
        Dim aCoWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value(aType & "_" & "COL", "DATAMSG") <> "", aIntPar.value(aType & "_" & "COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aPostClmn As String = If(aIntPar.value(aType & "_" & "COL", "DATAPOST") <> "", aIntPar.value(aType & "_" & "COL", "DATAPOST"), "INT-POST")
        Dim aPostClmnNr As Integer = 0
        Dim aCoClmnNr As Integer = If(aIntPar.value(aType & "_" & "COLNR", "DATACHECK") <> "", CInt(aIntPar.value(aType & "_" & "COLNR", "DATACHECK")), 1)
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value(aType & "_" & "RET", "OKMSG") <> "", aIntPar.value(aType & "_" & "RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.ThisAddIn.Application.ActiveWorkbook
        Try
            aCoWs = aWB.Worksheets(aCoWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aCoWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Co Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Co")
            Exit Sub
        End Try
        parseHeaderLine(aCoWs, jMax, pMsgClmn:=aMsgClmn, pMsgClmnNr:=aMsgClmnNr, pPostClmn:=aPostClmn, pPostClmnNr:=aPostClmnNr, pHdrLine:=aHdrLOff + 1)
        Try
            log.Debug("SapCoRibbon_AcctngManCostAlloc.Post - " & "processing data - disabling events, screen update, cursor")
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.ThisAddIn.Application.EnableEvents = False
            '            Globals.ThisAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aCoLOff + 1
            Dim aKey As String
            Dim aPost As String = ""
            Dim aItems As New TData(aIntPar)
            Dim aTSAP_Data_Co As New TSAP_Data_Co(aPar, aIntPar, aSAPAcctngMeta, "postManCostAlloc")
            Do
                If Left(CStr(aCoWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    If aPostClmnNr <> 0 Then
                        aPost = If(String.IsNullOrEmpty(CStr(aCoWs.Cells(i, aPostClmnNr).value)), "", CStr(aCoWs.Cells(i, aPostClmnNr).value))
                    End If
                    aKey = CStr(i)
                    ' read DATA
                    aItems.ws_parse_line_simple(aCoWs, aCoLOff, i, jMax, pHdrLine:=aHdrLOff + 1)
                    If String.IsNullOrEmpty(CStr(aCoWs.Cells(i + 1, aCoClmnNr).value)) Or aPost.ToUpper = "X" Then
                        If aTSAP_Data_Co.fillHeader(aItems) And aTSAP_Data_Co.fillData(aItems) Then
                            log.Debug("SapCoRibbon_AcctngManCostAlloc.Post - " & "calling aSAPAcctngManCostAlloc.Post")
                            Globals.ThisAddIn.Application.StatusBar = "Calling SAP-BAPI at line " & i
                            aRetStr = aSAPAcctngManCostAlloc.Post(aTSAP_Data_Co, pOKMsg:=aOKMsg, pCheck:=pCheck)
                            log.Debug("SapCoRibbon_AcctngManCostAlloc.Post - " & "aSAPAcctngManCostAlloc.Post returned, aRetStr=" & aRetStr)
                            For Each aKey In aItems.aTDataDic.Keys
                                aCoWs.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                            Next
                        End If
                        aItems = New TData(aIntPar)
                        aTSAP_Data_Co = New TSAP_Data_Co(aPar, aIntPar, aSAPAcctngMeta, "postManCostAlloc")
                    End If
                End If
                i += 1
            Loop While Not String.IsNullOrEmpty(CStr(aCoWs.Cells(i, aCoClmnNr).value))
            log.Debug("SapCoRibbon_AcctngManCostAlloc.Post - " & "all data processed - enabling events, screen update, cursor")
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapCoRibbon_AcctngManCostAlloc.Post failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP Co AddIn")
            log.Error("SapCoRibbon_AcctngManCostAlloc.Post - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Private Sub parseHeaderLine(ByRef pWs As Excel.Worksheet, ByRef pMaxJ As Integer, Optional pMsgClmn As String = "", Optional ByRef pMsgClmnNr As Integer = 0, Optional pHdrLine As Integer = 1, Optional pPostClmn As String = "", Optional ByRef pPostClmnNr As Integer = 0)
        pMaxJ = 0
        Do
            pMaxJ += 1
            If Not String.IsNullOrEmpty(pMsgClmn) And CStr(pWs.Cells(pHdrLine, pMaxJ).value) = pMsgClmn Then
                pMsgClmnNr = pMaxJ
            ElseIf Not String.IsNullOrEmpty(pPostClmn) And CStr(pWs.Cells(pHdrLine, pMaxJ).value) = pPostClmn Then
                pPostClmnNr = pMaxJ
            End If
        Loop While CStr(pWs.Cells(pHdrLine, pMaxJ + 1).value) <> ""
    End Sub

End Class
