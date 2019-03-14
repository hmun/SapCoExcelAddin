' Copyright 2016-2019 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class ConParamterRec
    Public aID As TField
    Public aName As TField
    Public aAppServerHost As TField
    Public aSystemNumber As TField
    Public aSystemID As TField
    Public aMessageServerHost As TField
    Public aLogonGroup As TField
    Public aTrace As TField
    Public aClient As TField
    Public aLanguage As TField
    Public aSncMode As TField
    Public aSncMyName As TField
    Public aSncPartnerName As TField

    Private sTField As TField

    Public Sub New()
        sTField = New TField

        aID = New TField
        aName = New TField
        aAppServerHost = New TField
        aSystemNumber = New TField
        aSystemID = New TField
        aMessageServerHost = New TField
        aLogonGroup = New TField
        aTrace = New TField
        aClient = New TField
        aLanguage = New TField
        aSncMode = New TField
        aSncMyName = New TField
        aSncPartnerName = New TField
    End Sub

    Public Function setValues(pID As String, pName As String, pAppServerHost As String, pSystemNumber As String, pSystemID As String, pMessageServerHost As String, pLogonGroup As String, pTrace As String, pClient As String, pLanguage As String, pSncMode As String, pSncMyName As String, pSncPartnerName As String)
        aID = sTField.create("ID", CStr(pID))
        aName = sTField.create("Name", CStr(pName))
        aAppServerHost = sTField.create("AppServerHost", CStr(pAppServerHost))
        aSystemNumber = sTField.create("SystemNumber", CStr(pSystemNumber))
        aSystemID = sTField.create("SystemID", CStr(pSystemID))
        aMessageServerHost = sTField.create("MessageServerHost", CStr(pMessageServerHost))
        aLogonGroup = sTField.create("LogonGroup", CStr(pLogonGroup))
        aTrace = sTField.create("Trace", CStr(pTrace))
        aClient = sTField.create("Client", CStr(pClient))
        aLanguage = sTField.create("Language", CStr(pLanguage))
        aSncMode = sTField.create("SncMode", CStr(pSncMode))
        aSncMyName = sTField.create("SncMyName", CStr(pSncMyName))
        aSncPartnerName = sTField.create("SncPartnerName", CStr(pSncPartnerName))
    End Function

    Public Function setValue(pField As String, pValue As String)
        If pField = "ID" Then
            aID = sTField.create(pField, CStr(pValue))
        ElseIf pField = "Name" Then
            aName = sTField.create(pField, CStr(pValue))
        ElseIf pField = "AppServerHost" Then
            aAppServerHost = sTField.create(pField, CStr(pValue))
        ElseIf pField = "SystemNumber" Then
            aSystemNumber = sTField.create(pField, CStr(pValue))
        ElseIf pField = "SystemID" Then
            aSystemID = sTField.create(pField, CStr(pValue))
        ElseIf pField = "MessageServerHost" Then
            aMessageServerHost = sTField.create(pField, CStr(pValue))
        ElseIf pField = "LogonGroup" Then
            aLogonGroup = sTField.create(pField, CStr(pValue))
        ElseIf pField = "Trace" Then
            aTrace = sTField.create(pField, CStr(pValue))
        ElseIf pField = "Client" Then
            aClient = sTField.create(pField, CStr(pValue))
        ElseIf pField = "Language" Then
            aLanguage = sTField.create(pField, CStr(pValue))
        ElseIf pField = "SncMode" Then
            aSncMode = sTField.create(pField, CStr(pValue))
        ElseIf pField = "SncMyName" Then
            aSncMyName = sTField.create(pField, CStr(pValue))
        ElseIf pField = "SncPartnerName" Then
            aSncPartnerName = sTField.create(pField, CStr(pValue))
        End If
    End Function

    Public Function getKey() As String
        Dim aKey As String
        aKey = aID.Value
        getKey = aKey
    End Function

    Public Function getRKey() As String
        Dim aKey As String
        aKey = aID.Value
        getRKey = aKey
    End Function

End Class

Public Class ConParameter
    Public aConCol As Collection
    Private sTField As TField

    Public Sub New()
        sTField = New TField
        aConCol = New Collection
    End Sub

    Public Function addCon(pID As String, pName As String, pAppServerHost As String, pSystemNumber As String, pSystemID As String, pMessageServerHost As String, pLogonGroup As String, pTrace As String, pClient As String, pLanguage As String, pSncMode As String, pSncMyName As String, pSncPartnerName As String)
        Dim aConRec As ConParamterRec
        Dim aKey As String
        aKey = pID
        If contains(aConCol, aKey, "obj") Then
            aConRec = aConCol(aKey)
            aConRec.setValues(pID, pName, pAppServerHost, pSystemNumber, pSystemID, pMessageServerHost, pLogonGroup, pTrace, pClient, pLanguage, pSncMode, pSncMyName, pSncPartnerName)
        Else
            aConRec = New ConParamterRec
            aConRec.setValues(pID, pName, pAppServerHost, pSystemNumber, pSystemID, pMessageServerHost, pLogonGroup, pTrace, pClient, pLanguage, pSncMode, pSncMyName, pSncPartnerName)
            aConCol.Add(aConRec, aKey)
        End If
    End Function

    Public Function addConValue(pID As String, pField As String, pValue As String)
        Dim aConRec As ConParamterRec
        Dim aKey As String
        aKey = pID
        If contains(aConCol, aKey, "obj") Then
            aConRec = aConCol(aKey)
            aConRec.setValue(pField, pValue)
        Else
            aConRec = New ConParamterRec
            aConRec.setValue("ID", pID)
            aConRec.setValue(pField, pValue)
            aConCol.Add(aConRec, aKey)
        End If
    End Function

    Private Function contains(col As Collection, Key As String, Optional aType As String = "var") As Boolean
        Dim obj As Object
        Dim var As Object
        On Error GoTo err
        contains = True
        If aType = "obj" Then
            obj = col(Key)
        Else
            var = col(Key)
        End If
        Exit Function
err:
        contains = False
    End Function

End Class