' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SAPAcctngPrimCostsItem
    Public SEND_CCTR As String
    Public SENACTTYPE As String
    Public SEN_ORDER As String
    Public SEN_WBS_EL As String
    Public SEN_NETWRK As String
    Public SENOPERATN As String
    Public SEND_FUNCTION As String
    Public PERSON_NO As String
    Public COST_ELEM As String
    Public VALUE_TCUR As Double
    Public SEG_TEXT As String
    Public REC_CCTR As String
    Public REC_ORDER As String
    Public REC_WBS_EL As String
    Public REC_NETWRK As String
    Public RECOPERATN As String
    Public REC_FUNCTION As String
    Public TRANS_CURR As String

    Public Function create(pSEND_CCTR As String, pSENACTTYPE As String, pSEN_ORDER As String, pSEN_WBS_EL As String, pSEN_NETWRK As String,
                            pSENOPERATN As String, pSEND_FUNCTION As String,
                            pPERSON_NO As String, pCOST_ELEM As String, pVALUE_TCUR As Double, pSEG_TEXT As String,
                            pREC_CCTR As String, pREC_ORDER As String, pREC_WBS_EL As String, pREC_NETWRK As String,
                            pRECOPERATN As String, pREC_FUNCTION As String,
                            pTRANS_CURR As String)
        Dim aSAPAcctngPrimCostsItem As New SAPAcctngPrimCostsItem
        aSAPAcctngPrimCostsItem.SEND_CCTR = pSEND_CCTR
        aSAPAcctngPrimCostsItem.SENACTTYPE = pSENACTTYPE
        aSAPAcctngPrimCostsItem.SEN_ORDER = pSEN_ORDER
        aSAPAcctngPrimCostsItem.SEN_WBS_EL = pSEN_WBS_EL
        aSAPAcctngPrimCostsItem.SEN_NETWRK = pSEN_NETWRK
        aSAPAcctngPrimCostsItem.SENOPERATN = pSENOPERATN
        aSAPAcctngPrimCostsItem.SEND_FUNCTION = pSEND_FUNCTION
        aSAPAcctngPrimCostsItem.PERSON_NO = pPERSON_NO
        aSAPAcctngPrimCostsItem.COST_ELEM = pCOST_ELEM
        aSAPAcctngPrimCostsItem.VALUE_TCUR = pVALUE_TCUR
        aSAPAcctngPrimCostsItem.SEG_TEXT = pSEG_TEXT
        aSAPAcctngPrimCostsItem.REC_CCTR = pREC_CCTR
        aSAPAcctngPrimCostsItem.REC_ORDER = pREC_ORDER
        aSAPAcctngPrimCostsItem.REC_WBS_EL = pREC_WBS_EL
        aSAPAcctngPrimCostsItem.REC_NETWRK = pREC_NETWRK
        aSAPAcctngPrimCostsItem.RECOPERATN = pRECOPERATN
        aSAPAcctngPrimCostsItem.REC_FUNCTION = pREC_FUNCTION
        aSAPAcctngPrimCostsItem.TRANS_CURR = pTRANS_CURR
        create = aSAPAcctngPrimCostsItem
    End Function

End Class
