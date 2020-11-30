'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Public Class csBankCheque
    'Public BC_ref As Integer
    'Public ChqNo As String
    'Public Bank As String
    'Public ChqDate As Date
    'Public Amount As Double

    'Public VouType As String
    'Public VouNo As String
    Public VouType As String
    Public SrcLedgerID As Integer
    Public DstLedgerID As Integer

    Public CleardDate As Date
    Public DepositDate As Date
    Public isDeposit As Boolean

    'Public MatchedAmount As Double
    Public ChqStatus As Boolean
    Public Comment As String
    Public isAlert As Boolean

    Public Function DT_ChequeDetailsTemplate() As DataTable
        Dim DT_ChequeDetails As New DataTable
        DT_ChequeDetails.Columns.Add(New DataColumn("Slno", GetType(Integer)))
        DT_ChequeDetails.Columns.Add(New DataColumn("BC_Ref1", GetType(Integer)))
        DT_ChequeDetails.Columns.Add(New DataColumn("ChequeNo"))
        DT_ChequeDetails.Columns.Add(New DataColumn("Bank"))
        DT_ChequeDetails.Columns.Add(New DataColumn("ChequeDate", GetType(Date)))
        DT_ChequeDetails.Columns.Add(New DataColumn("Amount", GetType(Decimal)))
        DT_ChequeDetails.Columns.Add(New DataColumn("Comment"))
        DT_ChequeDetails.Columns.Add(New DataColumn("Status"))
        DT_ChequeDetails.Columns.Add(New DataColumn("DstLedgerID"))
        DT_ChequeDetails.Columns.Add(New DataColumn("IsDeposit"))

        DT_ChequeDetails.Columns("Amount").DefaultValue = 0
        DT_ChequeDetails.Columns("BC_ref1").DefaultValue = 0
        DT_ChequeDetails.Columns("Status").DefaultValue = "OPEN"

        DT_ChequeDetails.Columns("SlNo").AutoIncrement = True
        DT_ChequeDetails.Columns("SlNo").AutoIncrementStep = 1
        DT_ChequeDetails.Columns("SlNo").AutoIncrementSeed = 1
        Return DT_ChequeDetails
    End Function


End Class

Public Class csRVCheque
    Public objRV As New csRV
    Public objBankCheque As New csBankCheque
End Class

Public Class csPVCheque
    Public objPV As New csPV
    Public objBankCheque As New csBankCheque
End Class


Public Class csDTTemplate
    Public Shared Function DT_FromSource() As DataTable
        DT_FromSource = New DataTable
        DT_FromSource.Columns.Add(New DataColumn("VouNo"))
        DT_FromSource.Columns.Add(New DataColumn("NetAmt", System.Type.GetType("System.Decimal")))
        DT_FromSource.Columns.Add(New DataColumn("RcvdAmt", System.Type.GetType("System.Decimal")))
        DT_FromSource.Columns.Add(New DataColumn("PDCAmt", System.Type.GetType("System.Decimal")))
        DT_FromSource.Columns.Add(New DataColumn("BalAmt", System.Type.GetType("System.Decimal")))
        Return DT_FromSource
    End Function

    Public Shared Function DT_VouMatching() As DataTable
        DT_VouMatching = New DataTable
        DT_VouMatching.Columns.Add(New DataColumn("SlNo", GetType(Integer)))
        DT_VouMatching.Columns.Add(New DataColumn("BC_Ref", GetType(Integer)))
        DT_VouMatching.Columns.Add(New DataColumn("ChequeNo"))
        DT_VouMatching.Columns.Add(New DataColumn("Date_", GetType(Date)))
        DT_VouMatching.Columns.Add(New DataColumn("Voucher"))
        DT_VouMatching.Columns.Add(New DataColumn("VouRef"))
        DT_VouMatching.Columns.Add(New DataColumn("PayType"))
        DT_VouMatching.Columns.Add(New DataColumn("VouType"))
        DT_VouMatching.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Decimal")))
        DT_VouMatching.Columns.Add(New DataColumn("PaidAmt", System.Type.GetType("System.Decimal")))
        DT_VouMatching.Columns.Add(New DataColumn("PDCAmt", System.Type.GetType("System.Decimal")))
        DT_VouMatching.Columns.Add(New DataColumn("BalAmt", System.Type.GetType("System.Decimal")))
        'DT_VouMatching.Columns.Add(New DataColumn("MatchThisRV", System.Type.GetType("System.Double")))
        DT_VouMatching.Columns.Add(New DataColumn("PayNow", System.Type.GetType("System.Decimal")))
        DT_VouMatching.Columns.Add(New DataColumn("FullPay", GetType(Boolean)))
        DT_VouMatching.Columns.Add(New DataColumn("RefNo"))
        DT_VouMatching.Columns("BC_Ref").DefaultValue = 0
        Return DT_VouMatching
    End Function

    Public Shared Function DT_4Dialog() As DataTable
        DT_4Dialog = New DataTable
        DT_4Dialog.Columns.Add(New DataColumn("SlNo", GetType(Integer)))
        DT_4Dialog.Columns.Add(New DataColumn("BC_Ref", GetType(Integer)))
        DT_4Dialog.Columns.Add(New DataColumn("Date_", GetType(Date)))
        DT_4Dialog.Columns.Add(New DataColumn("ChequeNo"))
        DT_4Dialog.Columns.Add(New DataColumn("Voucher"))
        DT_4Dialog.Columns.Add(New DataColumn("VouRef"))
        DT_4Dialog.Columns.Add(New DataColumn("PayType"))
        DT_4Dialog.Columns.Add(New DataColumn("VouType"))
        DT_4Dialog.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Decimal")))
        DT_4Dialog.Columns.Add(New DataColumn("PaidAmt", System.Type.GetType("System.Decimal")))
        DT_4Dialog.Columns.Add(New DataColumn("PDCAmt", System.Type.GetType("System.Decimal")))
        DT_4Dialog.Columns.Add(New DataColumn("BalAmt", System.Type.GetType("System.Decimal")))
        'DT_4Dialog.Columns.Add(New DataColumn("MatchThisRV", System.Type.GetType("System.Double")))
        DT_4Dialog.Columns.Add(New DataColumn("PayNow", System.Type.GetType("System.Decimal")))
        'DT_4Dialog.Columns.Add(New DataColumn("FullPay", GetType(Boolean)))
        'DT_4Dialog.Columns.Add(New DataColumn("RefNo"))
        DT_4Dialog.Columns("PayType").DefaultValue = False
        DT_4Dialog.Columns("ChequeNo").DefaultValue = False
        Return DT_4Dialog
    End Function
End Class