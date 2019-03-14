' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SAPFormat

    Public Function unpack(val As String, length As Integer) As String
        Dim ZeroStr As String
        If IsNumeric(val) Then
            ZeroStr = "000000000000000000000000000000"
            unpack = Left(ZeroStr, length - Len(val)) & val
        Else
            unpack = val
        End If
    End Function

    Public Function pspid(val As String, length As Integer) As String
        Dim ZeroStr As String
        ZeroStr = "000000000000000000000000000000"
        val = Replace(val, ".", "")
        pspid = val & Left(ZeroStr, length - Len(val))
    End Function

End Class
