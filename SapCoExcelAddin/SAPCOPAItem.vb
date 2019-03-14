Public Class SAPCOPAItem
    Public gFIELDNAME As String
    Public gVALUE As Object
    Public gCURRENCY As String

    Public Function create(pFIELDNAME As String, pVALUE As Object, pCURRENCY As String) As SAPCOPAItem
        Dim aSAPCOPAItem As New SAPCOPAItem
        aSAPCOPAItem.gFIELDNAME = pFIELDNAME
        aSAPCOPAItem.gVALUE = pVALUE
        aSAPCOPAItem.gCURRENCY = pCURRENCY
        create = aSAPCOPAItem
    End Function
End Class
