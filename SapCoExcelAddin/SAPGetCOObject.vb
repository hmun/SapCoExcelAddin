' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPGetCOObject
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        Try
            sapcon = aSapCon
            destination = aSapCon.getDestination()
            sapcon.checkCon()
        Catch ex As System.Exception
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPGetCOObject")
        End Try
    End Sub

    Public Function GetCoObjects(pType As String, pFiscy As String, pVersn As String,
                                 pKokrs As String, pBukrs As String, pObjects As Collection) As String
        GetCoObjects = "Failed"
        Try
            oRfcFunction = destination.Repository.CreateFunction("ZCOPC_GET_COOBJ")
            Dim oObjects As IRfcTable = oRfcFunction.GetTable("T_OBJECTS")
            oObjects.Clear()
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            oRfcFunction.SetValue("I_TYPE", pType)
            oRfcFunction.SetValue("I_VERSN", lSAPFormat.unpack(pVersn, 3))
            oRfcFunction.SetValue("I_GJAHR", pFiscy)
            oRfcFunction.SetValue("I_KOKRS", pKokrs)
            If pBukrs <> "" Then
                oRfcFunction.SetValue("I_BURKS", pBukrs)
            End If
            ' call the Function
            oRfcFunction.Invoke(destination)
            For i As Integer = 0 To oObjects.Count - 1
                Dim lSAPCOObject As New SAPCOObject
                If Not (pType = "I" And oObjects(i).GetValue("SKOSTL") = "") Then
                    lSAPCOObject = lSAPCOObject.create(oObjects(i).GetValue("KOSTL"),
                                                       oObjects(i).GetValue("LSTAR"),
                                                       oObjects(i).GetValue("KSTAR"),
                                                       oObjects(i).GetValue("SKOSTL"),
                                                       oObjects(i).GetValue("SLSTAR"),
                                                       oObjects(i).GetValue("WBS_ELEMENT"))
                    pObjects.Add(lSAPCOObject)
                End If
            Next i
            If pObjects.Count <> 0 Then
                GetCoObjects = "Success"
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPGetCOObject")
            GetCoObjects = "Error: Exception in GetCoObjects"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function
End Class
