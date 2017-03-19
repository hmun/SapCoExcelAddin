' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SAPAcctngActivityItem
    Public SEND_CCTR As String
    Public PERSON_NO As String
    Public ACTTYPE As String
    Public ACTVTY_QTY As Double
    Public SEG_TEXT As String
    Public REC_WBS_EL As String
    Public REC_NETWRK As String
    Public RECOPERATN As String
    Public REC_ORDER As String
    Public REC_CCTR As String
    Public REC_FUNCTION As String
    Public PRICE As Double
    Public PRICE_FIX As Double
    Public PRICE_VAR As Double
    Public PRICE_UNIT As Integer
    Public CURR As String

    Public Function create(pSEND_CCTR As String, pPERSON_NO As String, pACTTYPE As String, pACTVTY_QTY As Double, pSEG_TEXT As String,
                            pREC_WBS_EL As String, pREC_NETWRK As String, pRECOPERATN As String, pREC_ORDER As String, pREC_CCTR As String,
                            pPRICE As Double, pPRICE_FIX As Double, pPRICE_VAR As Double, pPRICE_UNIT As Integer,
                            pCURR As String, Optional ByVal pREC_FUNCTION As String = "") As SAPAcctngActivityItem
        Dim aSAPAcctngActivityItem As New SAPAcctngActivityItem
        aSAPAcctngActivityItem.SEND_CCTR = pSEND_CCTR
        aSAPAcctngActivityItem.PERSON_NO = pPERSON_NO
        aSAPAcctngActivityItem.ACTTYPE = pACTTYPE
        aSAPAcctngActivityItem.ACTVTY_QTY = pACTVTY_QTY
        aSAPAcctngActivityItem.SEG_TEXT = pSEG_TEXT
        aSAPAcctngActivityItem.REC_WBS_EL = pREC_WBS_EL
        aSAPAcctngActivityItem.REC_NETWRK = pREC_NETWRK
        aSAPAcctngActivityItem.RECOPERATN = pRECOPERATN
        aSAPAcctngActivityItem.REC_ORDER = pREC_ORDER
        aSAPAcctngActivityItem.REC_CCTR = pREC_CCTR
        aSAPAcctngActivityItem.REC_FUNCTION = pREC_FUNCTION
        aSAPAcctngActivityItem.PRICE = pPRICE
        aSAPAcctngActivityItem.PRICE_VAR = pPRICE_VAR
        aSAPAcctngActivityItem.PRICE_FIX = pPRICE_FIX
        aSAPAcctngActivityItem.PRICE_UNIT = pPRICE_UNIT
        aSAPAcctngActivityItem.CURR = pCURR
        create = aSAPAcctngActivityItem
    End Function

End Class
