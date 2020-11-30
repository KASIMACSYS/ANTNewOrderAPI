
Public Class csCRDR
    Inherits csSignature

    Public int_CID As Integer
    Public objCRDRMain As New csCRDRMain
    Public objCRDRSub As New csCRDRSub
    Public objProject As New csProjectDetail
   
    
    Public Function DBTemplate() As DataTable
        Dim DT_Template As New DataTable
        DT_Template.Columns.Add(New DataColumn("SlNo", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("Ledger"))
        DT_Template.Columns.Add(New DataColumn("DstLedgerID", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Comment"))
        DT_Template.Columns("Amount").DefaultValue = 0
        DT_Template.Columns("SlNo").AutoIncrement = True
        DT_Template.Columns("SlNo").AutoIncrementStep = 1
        DT_Template.Columns("SlNo").AutoIncrementSeed = 1
        Return DT_Template
    End Function

    Public Function VouMatching() As DataTable
        Dim dt_VouMatching As New DataTable
        dt_VouMatching.Columns.Add(New DataColumn("SlNo", GetType(Integer)))
        dt_VouMatching.Columns.Add(New DataColumn("BC_Ref", GetType(Integer)))
        dt_VouMatching.Columns.Add(New DataColumn("ChequeNo"))
        dt_VouMatching.Columns.Add(New DataColumn("Voucher"))
        dt_VouMatching.Columns.Add(New DataColumn("VouRef"))
        dt_VouMatching.Columns.Add(New DataColumn("PayType"))
        dt_VouMatching.Columns.Add(New DataColumn("PDCType"))
        dt_VouMatching.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Double")))
        dt_VouMatching.Columns.Add(New DataColumn("RcvdAmt", System.Type.GetType("System.Double")))
        dt_VouMatching.Columns.Add(New DataColumn("PDCAmt", System.Type.GetType("System.Double")))
        dt_VouMatching.Columns.Add(New DataColumn("BalAmt", System.Type.GetType("System.Double")))
        dt_VouMatching.Columns.Add(New DataColumn("PayNow", System.Type.GetType("System.Double")))
        dt_VouMatching.Columns.Add(New DataColumn("FullPay", GetType(Boolean)))
        dt_VouMatching.Columns.Add(New DataColumn("RefNo"))
        Return dt_VouMatching
    End Function
    

    Public Function InvoiceAccountTemplate() As DataTable
        Dim DTInvoiceAccount As New DataTable
        DTInvoiceAccount.Columns.Add("SlNo", GetType(Integer))
        DTInvoiceAccount.Columns.Add("Ledger")
        DTInvoiceAccount.Columns.Add("DstLedgerID", GetType(Integer))
        DTInvoiceAccount.Columns.Add("Amount", System.Type.GetType("System.Double"))
        DTInvoiceAccount.Columns.Add(New DataColumn("Tax"))
        DTInvoiceAccount.Columns.Add(New DataColumn("TaxPercentage", System.Type.GetType("System.Double")))
        DTInvoiceAccount.Columns.Add(New DataColumn("TaxAmount", System.Type.GetType("System.Decimal")))
        DTInvoiceAccount.Columns.Add(New DataColumn("NonClaimableTaxAmount", System.Type.GetType("System.Decimal")))
        DTInvoiceAccount.Columns.Add("NetAmount", System.Type.GetType("System.Double"))
        DTInvoiceAccount.Columns.Add("TRNNo")
        DTInvoiceAccount.Columns.Add("Comment")
        DTInvoiceAccount.Columns.Add(New DataColumn("ItemTaxDetails"))
        Return DTInvoiceAccount
    End Function
End Class

Public Class csCRDRMain
    Public str_Flag As String
    Public str_MenuID As String

    Public int_BusinessPeriodID As Integer
    Public str_VouNo As String
    Public int_RevNo As Integer
    Public str_Prefix As String
    Public str_Type As String
    Public str_SrcLedgerID As Integer
    Public str_Alias As String
    Public dtp_VouDate As Date
    Public str_Comment As String
    Public dbl_TCTotalAmount As Double
    Public dbl_TCDisAmount As Double
    Public dbl_TCDiscountAmount As Double
    Public dbl_TCNetAmount As Double
    Public dbl_LCNetAmount As Double
    Public dbl_TCPaidAmount As Double
    Public dbl_TCPDCAmount As Double
    Public str_TCCurrency As String
    Public dbl_ExchangeRate As Double
    Public str_VouRefNo As String
    Public bool_TaxFileReturn As Boolean
    Public bool_PaymentType As Boolean
    Public str_CashorCredit As String
    Public int_CashLedger As Integer
    Public dbl_CashTendered As Double
    Public str_ItemTaxCode As String
    Public str_InvoiceTaxCode As String
    Public str_InvoiceTaxXML As String
    Public dbl_TCItemTaxAmount As Double
    Public dbl_TCInvoiceTaxAmount As Double
    Public dt_TaxItemDetails As DataTable
End Class

Public Class csCRDRSub
    Public dt_DstLedgerDetails As DataTable
    Public dt_InvMatching As DataTable
End Class

