' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPZ_BC_EXCEL_ADDIN_VERS_CHK

    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        sapcon = aSapCon
        destination = aSapCon.getDestination()
        Try
            oRfcFunction = destination.Repository.CreateFunction("Z_BC_EXCEL_ADDIN_VERS_CHK")
        Catch ex As Exception
            oRfcFunction = Nothing
        End Try
    End Sub

    Public Function checkVersion(pAddIn As String, pVersion As String) As Integer
        sapcon.checkCon()
        If oRfcFunction Is Nothing Then
            ' for systems that do not contain Z_BC_EXCEL_ADDIN_VERS_CHK we can not check the version
            checkVersion = 0
        Else
            Try
                Dim oRETURN As IRfcTable = oRfcFunction.GetTable("T_RETURN")
                Dim oE_ALLOWED_VERSION As IRfcStructure = oRfcFunction.GetStructure("E_ALLOWED_VERSION")
                oRETURN.Clear()

                oRfcFunction.SetValue("I_ADDIN", pAddIn)
                oRfcFunction.SetValue("I_VERSION", pVersion)

                oRfcFunction.Invoke(destination)
                If oRETURN.Count > 0 Then
                    If oRETURN(0).GetValue("TYPE") = "S" Then
                        checkVersion = 0
                    Else
                        checkVersion = 4
                    End If
                Else
                    checkVersion = 8
                End If
            Catch abap_ex As RfcAbapBaseException
                Select Case abap_ex.Message
                    Case "WRONG_VERSION_FORMAT"
                        checkVersion = 1
                    Case "UNSUPPORTED_VERSION"
                        checkVersion = 2
                    Case "NO_VERSION_MAINTAINED"
                        checkVersion = 3
                    Case Else
                        checkVersion = 8
                End Select
                Exit Function
            Catch ex As Exception
                MsgBox("Exception in checkVersion! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPZ_BC_EXCEL_ADDIN_VERS_CHK")
                checkVersion = 8
            End Try
        End If
        Exit Function
    End Function

End Class
