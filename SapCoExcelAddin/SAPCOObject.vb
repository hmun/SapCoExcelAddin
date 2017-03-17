' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SAPCOObject
    Public Costcenter As String
    Public Acttype As String
    Public Costelem As String
    Public SCostcenter As String
    Public SActtype As String
    Public WBS_ELEMENT As String
    Public STATKEYFIG As String

    Public Function create(pCostcenter As String, pActtype As String, pCostelem As String,
                            Optional pSCostcenter As String = "",
                            Optional pSActtype As String = "",
                            Optional pWBS_ELEMENT As String = "",
                            Optional pSTATKEYFIG As String = "") As SAPCOObject
        Dim aSAPCOObject As New SAPCOObject

        aSAPCOObject.Costcenter = pCostcenter
        aSAPCOObject.Acttype = pActtype
        aSAPCOObject.Costelem = pCostelem
        aSAPCOObject.SCostcenter = pSCostcenter
        aSAPCOObject.SActtype = pSActtype
        aSAPCOObject.WBS_ELEMENT = pWBS_ELEMENT
        aSAPCOObject.STATKEYFIG = pSTATKEYFIG
        create = aSAPCOObject
    End Function

End Class
