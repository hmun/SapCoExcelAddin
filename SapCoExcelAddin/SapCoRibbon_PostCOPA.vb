Imports SAPLogon

Public Class SapCoRibbon_PostCOPA
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private aOperatingConcern As String
    Private aMaxLines As String

    Private Function getCOPAParameters() As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim aIntPar As New SAPCommon.TStr
        Dim aPwsName As String = "Parameter"
        ' get internal parameters
        If getIntParameters(aIntPar) Then
            aPwsName = If(aIntPar.value("WS", "PARA_PA") <> "", aIntPar.value("WS", "PARA_PA"), "Parameter")
        End If
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets(aPwsName)
        Catch Exc As System.Exception
            MsgBox("No " & aPwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP CO-PA Actuals Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getCOPAParameters = False
            Exit Function
        End Try
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> "SAPCostingBasedData" And aKey <> "SAPCoMultiple" Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key SAPCostingBasedData. Check if the current workbook is a valid SAP CO-PA Actuals Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getCOPAParameters = False
            Exit Function
        End If
        aOperatingConcern = CStr(aPws.Cells(2, 2).Value)
        aMaxLines = CInt(aPws.Cells(3, 2).Value)
        If aOperatingConcern = "" Then
            MsgBox("Please fill all obligatory fields in the parameter sheet!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap RepstPrimCosts")
            getCOPAParameters = False
            Exit Function
        End If
        getCOPAParameters = True
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
            '            MsgBox("No Parameter_Int Sheet in current workbook. Check if the current workbook is a valid SapCo Template",
            '           MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCo")
            '            getIntParameters = False
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
        Dim aSAPCOPAItem As New SAPCOPAItem
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        Dim aSAPFormat As New SAPCommon.SAPFormat
        Dim aSAPProjectDefinition As New SAPProjectDefinition(pSapCon)
        Dim aSAPWbsElement As New SAPWbsElement(pSapCon)
        Dim aData As New Collection
        Dim aDataRow As New Collection
        Dim aLines As Integer
        Dim aStartLine As Integer
        Dim aEndLine As Integer
        Dim aLineCnt As Integer

        Dim i As Integer
        Dim j As Integer
        Dim maxJ As Integer
        Dim aRetStr As String

        Dim aFIELDNAME As String
        Dim aVALUE As Object
        Dim aCURRENCY As String

        Dim aCells As Excel.Range
        Dim aIntPar As New SAPCommon.TStr
        Dim aDwsName As String = "Data"
        ' get internal parameters
        If getIntParameters(aIntPar) Then
            aDwsName = If(aIntPar.value("WS", "DATA_PA") <> "", aIntPar.value("WS", "DATA_PA"), "Data")
        End If
        If getCOPAParameters() = False Then
            Exit Sub
        End If
        Dim aSAPCOPAActuals As New SAPCOPAActuals(pSapCon, pIntPar:=aIntPar)
        aWB = Globals.SapCoExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets(aDwsName)
        Catch Exc As System.Exception
            MsgBox("No " & aDwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP CO-PA Actuals Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-PA")
            Exit Sub
        End Try
        ' Read the Items
        aDws.Activate()
        Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
        Globals.SapCoExcelAddin.Application.EnableEvents = False
        Globals.SapCoExcelAddin.Application.ScreenUpdating = False
        i = 5
        ' determine the last column
        maxJ = 1
        Do
            maxJ = maxJ + 1
        Loop While CStr(aDws.Cells(1, maxJ).Value) <> ""
        aStartLine = i
        aLineCnt = 0
        aData = New Collection
        Do
            If Left(CStr(aDws.Cells(i, maxJ).Value), 7) <> "Success" Then
                aDataRow = New Collection
                j = 1
                Do
                    aVALUE = ""
                    aCURRENCY = ""
                    aFIELDNAME = ""
                    aSAPCOPAItem = New SAPCOPAItem
                    If aDws.Cells(2, j).Value IsNot Nothing Then
                        aCURRENCY = CStr(aDws.Cells(2, j).Value)
                        If aDws.Cells(i, j).Value IsNot Nothing Then
                            aVALUE = aSAPFormat.dec(CStr(aDws.Cells(i, j).Value), 2)
                        Else
                            aVALUE = aSAPFormat.dec("0", 2)
                        End If
                    Else
                        aCURRENCY = ""
                        If aDws.Cells(i, j).Value IsNot Nothing Then
                            Select Case CStr(aDws.Cells(3, j).Value)
                                Case "DATE"
                                    Try
                                        aVALUE = CDate(aDws.Cells(i, j).Value).ToString("yyyyMMdd")
                                    Catch Exc As System.Exception
                                        aVALUE = ""
                                    End Try
                                Case "PERIO"
                                    aVALUE = Right(aDws.Cells(i, j).Value, 4) & Left(aDws.Cells(i, j).Value, 3)
                                Case "PROJ"
                                    If CStr(aDws.Cells(i, j).Value) <> "" Then
                                        aVALUE = aSAPProjectDefinition.GetPspnr(CStr(aDws.Cells(i, j).Value))
                                    Else
                                        aVALUE = ""
                                    End If
                                Case "WBS"
                                    If CStr(aDws.Cells(i, j).Value) <> "" Then
                                        aVALUE = aSAPWbsElement.GetPspnr(CStr(aDws.Cells(i, j).Value))
                                    Else
                                        aVALUE = ""
                                    End If
                                Case Else
                                    If Left(aDws.Cells(3, j).Value, 1) = "U" Then
                                        aVALUE = aSAPFormat.unpack(aDws.Cells(i, j).Value, CInt(Right(aDws.Cells(3, j).Value, Len(aDws.Cells(3, j).Value) - 1)))
                                    ElseIf Left(aDws.Cells(3, j).Value, 1) = "P" Then
                                        aVALUE = aSAPFormat.pspid(aDws.Cells(i, j).Value, CInt(Right(aDws.Cells(3, j).Value, Len(aDws.Cells(3, j).Value) - 1)))
                                    Else
                                        aVALUE = CStr(aDws.Cells(i, j).Value)
                                    End If
                            End Select
                        End If
                    End If

                    aFIELDNAME = CStr(aDws.Cells(1, j).Value)
                    aSAPCOPAItem = aSAPCOPAItem.create(aFIELDNAME, aVALUE, aCURRENCY)
                    aDataRow.Add(aSAPCOPAItem)
                    j = j + 1
                Loop While CStr(aDws.Cells(1, j).Value) <> ""
                aData.Add(aDataRow)
                aLineCnt = aLineCnt + 1
                If CInt(aMaxLines) = 1 Then
                    '     post the line
                    Globals.SapCoExcelAddin.Application.StatusBar = "Posting at line " & i
                    aRetStr = aSAPCOPAActuals.PostCostingBasedData(aOperatingConcern, aData, pCheck:=pCheck)
                    aDws.Cells(i, j).Value = aRetStr
                    aData = New Collection
                ElseIf aLineCnt >= CInt(aMaxLines) Then
                    aEndLine = i
                    '     post the lines
                    Globals.SapCoExcelAddin.Application.StatusBar = "Posting at line " & aEndLine
                    aRetStr = aSAPCOPAActuals.PostCostingBasedData(aOperatingConcern, aData, pCheck:=pCheck)
                    aCells = aDws.Range(aDws.Cells(aStartLine, j), aDws.Cells(aEndLine, j))
                    aCells.Value = aRetStr
                    aStartLine = i + 1
                    aLineCnt = 0
                    aData = New Collection
                End If
            Else
                aDws.Cells(i, maxJ + 1).Value = "ignored - already posted"
            End If
            i = i + 1
        Loop While CStr(aDws.Cells(i, 1).Value) <> ""
        ' post the rest
        If aData.Count > 0 Then
            aEndLine = i - 1
            Globals.SapCoExcelAddin.Application.StatusBar = "Posting at line " & aEndLine
            aRetStr = aSAPCOPAActuals.PostCostingBasedData(aOperatingConcern, aData, pCheck:=pCheck)
            aCells = aDws.Range(aDws.Cells(aStartLine, j), aDws.Cells(aEndLine, j))
            aCells.Value = aRetStr
        End If
        Globals.SapCoExcelAddin.Application.EnableEvents = True
        Globals.SapCoExcelAddin.Application.ScreenUpdating = True
        Globals.SapCoExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
    End Sub
End Class
