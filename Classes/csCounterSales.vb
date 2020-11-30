'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class csSalesInvoice
    Inherits csSignature

    Public int_CID As String
    Public objSalInvMain As New csSalesInvoiceMain
    Public objSalInvSub As New csSalesInvoiceSub
    Public objMerchantDetails As New csCustomerDetails
    Public objProject As csProjectDetail
    Public objRetention As New csRetention
    Public DTBatch As New DataTable
    Public DTItemExtraDetails As DataTable

    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        '' If CustomerSetting.Item("Project").ToString = "True" Then
        objProject = New csProjectDetail
        '' End If
    End Sub

    Public Function DBTemplate() As DataTable
        Dim DT_Template As New DataTable
        DT_Template.Columns.Add(New DataColumn("SortNo", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("BarCodeNo"))
        DT_Template.Columns.Add(New DataColumn("Alias1"))
        DT_Template.Columns.Add(New DataColumn("Alias2"))
        DT_Template.Columns.Add(New DataColumn("ItemCode"))
        DT_Template.Columns.Add(New DataColumn("Unit"))
        DT_Template.Columns.Add(New DataColumn("BaseUnit", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("VouQty", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("PrimaryQty", System.Type.GetType("System.Double")))
 		DT_Template.Columns.Add(New DataColumn("PriceType"))
        DT_Template.Columns.Add(New DataColumn("Package", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Pieces", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Price", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("BaseUnitPrice", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("DiscType"))
        DT_Template.Columns.Add(New DataColumn("DiscPercentage", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("TCDiscountAmount", System.Type.GetType("System.Decimal")))
        DT_Template.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Decimal")))
        DT_Template.Columns.Add(New DataColumn("Tax"))
        DT_Template.Columns.Add(New DataColumn("TaxPercentage", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("TaxAmount", System.Type.GetType("System.Decimal")))
        DT_Template.Columns.Add(New DataColumn("NonClaimableTaxAmount", System.Type.GetType("System.Decimal")))
        DT_Template.Columns.Add(New DataColumn("NetAmount", System.Type.GetType("System.Decimal")))
        DT_Template.Columns.Add(New DataColumn("LCAmount", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("LCCostPrice", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("LCCostAmount", System.Type.GetType("System.Decimal")))
        DT_Template.Columns.Add(New DataColumn("BalanceQty", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("DeliveredTotQty", System.Type.GetType("System.Double")))
       
        DT_Template.Columns.Add(New DataColumn("Comment"))
        DT_Template.Columns.Add(New DataColumn("PartNo"))
        DT_Template.Columns.Add(New DataColumn("Desc1"))
        DT_Template.Columns.Add(New DataColumn("Desc2"))
        DT_Template.Columns.Add(New DataColumn("Desc3"))
        DT_Template.Columns.Add(New DataColumn("Desc4"))
        DT_Template.Columns.Add(New DataColumn("Desc5"))
        DT_Template.Columns.Add(New DataColumn("Desc6"))
        DT_Template.Columns.Add(New DataColumn("Desc7"))
        DT_Template.Columns.Add(New DataColumn("Desc8"))
        DT_Template.Columns.Add(New DataColumn("OrgSlno", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("WHID"))
        DT_Template.Columns.Add(New DataColumn("Size1", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Size2", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Size3", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("MinSellPrice", System.Type.GetType("System.Decimal")))
        DT_Template.Columns.Add(New DataColumn("ItemTaxDetails"))
        Return DT_Template
    End Function

    ''Public Function InvoiceAccountTemplate() As DataTable
    ''    Dim DTInvoiceAccount As New DataTable

    ''    DTInvoiceAccount.Columns.Add("VouRef", GetType(Integer))
    ''    DTInvoiceAccount.Columns.Add("LedgerID", GetType(Integer))
    ''    DTInvoiceAccount.Columns.Add("TCDebit", System.Type.GetType("System.Double"))
    ''    DTInvoiceAccount.Columns.Add("TCCredit", System.Type.GetType("System.Double"))
    ''    DTInvoiceAccount.Columns.Add("Comment")
    ''    Return DTInvoiceAccount
    ''End Function

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
        Return dt_VouMatching
    End Function
End Class

Public Class csSalesInvoiceMain
    Public str_Flag As String
    Public str_MenuID As String
    Public str_Prefix As String
    Public int_SeqNo As Integer
    Public str_SalInvNo As String
    Public str_LpoNo As String
    Public str_DONo As String
    Public str_SalOrd As String
    Public str_InvRef As String
    Public int_RevNo As Integer
    Public str_SrcLedgerID As String
    Public str_Alias As String
    Public dtp_InvDate As Date
    Public dtp_DueDate As Date
    Public int_Aging As Integer
    Public str_PayTerm As String
    Public str_SalesManID As String
    Public str_InvoiceType As String
    Public bool_SalesInvoice As Boolean
    Public bool_IsCashSales As Boolean
    Public bool_AffectInventory As Boolean
    Public str_PaymentStatus As String
    Public str_InvoiceStatus As String
    Public str_Comment As String
    Public str_DeliveryAddress As String

    Public dbl_TCAmount As Double
    Public dbl_TCTotalAmount As Double
    Public dbl_TCDisAmount As Double
    Public dbl_TCDiscountAmount As Double
    Public dbl_TCMiscPercentage As Double
    Public dbl_TCMiscAmount As Double
    Public dbl_TCAdjAmount As Double
    Public dbl_TCNetAmount As Double

    Public dbl_LCNetCostPrice As Double
    Public dbl_LCNetProfit As Double

    Public dbl_TCPDCAmount As Double

    Public dbl_LCNetAmount As Double
    Public dbl_LCPDCAmount As Double
    Public str_CurrencyID As String
    Public dbl_ExchangeRate As Double
    Public str_MiscText As String
    Public str_DiscText As String
    Public int_BusinessPeriodID As Integer
    Public str_CashorCredit As String
    Public int_CashLedger As Integer
    Public dbl_CashTendered As Double
    Public int_StatusCancel As Integer
    Public int_LanguageCode As Integer

    Public dt_DONo4SIS As DataTable
    Public dt_InvoiceAccounts As DataTable
    Public dt_TaxItemDetails As DataTable

    Public str_WHID As String

    ''Air Master Specific
    Public str_Desc1 As String
    Public str_Desc2 As String
    Public str_Desc3 As String
    Public str_Desc4 As String
    Public str_Desc5 As String
    Public str_Desc6 As String
    Public str_Desc7 As String
    Public str_Desc8 As String
    Public str_Desc9 As String
    Public str_Desc10 As String

    Public str_ContactPerson As String

    Public str_UserComment As String = String.Empty
    Public str_ApproverComment As String = String.Empty
    Public bool_TaxFileReturn As Boolean
    Public str_ItemTaxCode As String
    Public str_InvoiceTaxCode As String
    Public str_InvoiceTaxXML As String
    Public dbl_TCItemTaxAmount As Double
    Public dbl_TCInvoiceTaxAmount As Double
    Public dbl_ItemDiscPercentage As Double
    Public str_Country As String
    Public str_Filter3 As String
    Public str_Filter4 As String
    Public str_Consignee As String

    Public str_RTF_Description As String
    Public int_RevisionHistoryNo As Integer
End Class

Public Class csSalesInvoiceSub
    Public dt_SalInv As DataTable
    Public dt_SalInvMatching As DataTable
End Class






