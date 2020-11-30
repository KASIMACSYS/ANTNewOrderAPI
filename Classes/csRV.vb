'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class csRV
    Inherits csSignature
    Public int_CID As Integer
    Public int_BusinessPeriodID As Integer
    Public str_Flag As String
    Public str_MenuID As String
    Public str_FormPrefix As String

    Public str_VouNo As String
    Public str_VouRef As String
    Public int_RevNo As Integer
    Public str_LedgerType As String
    Public dtp_RVDate As Date
    Public str_SrcLedgerID As String '= String.Empty
    Public str_Alice As String
    Public str_DstLedgerID As String
    Public str_PayType As String
    Public str_BCRef As String
    Public str_Comment As String

    Public dbl_TCTotalAmount As Decimal
    Public dbl_TCDisAmount As Decimal
    Public dbl_TCDiscountAmount As Decimal
    Public dbl_TCMiscAmount As Decimal
    Public dbl_TCNetAmount As Decimal
    Public dbl_LCNetAmount As Decimal
    Public str_CurrencyID As String
    Public dbl_ExchangeRate As Double
    Public str_CashLedger As String
    Public str_VouForHRMode As String

    Public int_StatusCancel As Integer
    Public int_StatusCancelPrevious As Integer

    Public str_Desc1 As String
    Public str_Desc2 As String
    Public str_Desc3 As String
    Public str_Desc4 As String
    Public str_Desc5 As String
    Public str_Desc6 As String
    Public str_Desc7 As String
    Public str_Desc8 As String
    'Public str_UserComment As String = String.Empty
    Public int_LanguageCode As Integer

    Public dt_RVMatching As DataTable
    Public DT_Wages As DataTable

    Public Function DT_VouMatching() As DataTable
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
        DT_VouMatching.Columns.Add(New DataColumn("PayNow", System.Type.GetType("System.Decimal")))
        DT_VouMatching.Columns.Add(New DataColumn("FullPay", GetType(Boolean)))
        DT_VouMatching.Columns.Add(New DataColumn("RefNo"))
        DT_VouMatching.Columns("BC_Ref").DefaultValue = 0
        Return DT_VouMatching
    End Function
End Class

