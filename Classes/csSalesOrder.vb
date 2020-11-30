'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Public Class csSalesOrder
    Inherits csSignature

    Public int_CID As Integer
    Public objSalesOrderMain As New csSalesOrderMain
    Public objSalesorderSub As New csSalesOrderSub
    Public objMerchantDetails As New csCustomerDetails
    Public objproject As csProjectDetail
    Public DTBatch As DataTable
    Public DTItemExtraDetails As DataTable

    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        ''If CustomerSetting.Item("Project").ToString = "True" Then
        objproject = New csProjectDetail
        '' End If
    End Sub

    Public Function DBTemplate() As DataTable
        Dim DT_SOTemplate As New DataTable
        DT_SOTemplate.Columns.Add(New DataColumn("SortNo", System.Type.GetType("System.Int32")))
        DT_SOTemplate.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
        DT_SOTemplate.Columns.Add(New DataColumn("BarCodeNo"))
        DT_SOTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_SOTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_SOTemplate.Columns.Add(New DataColumn("ItemCode"))
        DT_SOTemplate.Columns.Add(New DataColumn("Package", System.Type.GetType("System.Double")))
        DT_SOTemplate.Columns.Add(New DataColumn("Pieces", System.Type.GetType("System.Double")))
        DT_SOTemplate.Columns.Add(New DataColumn("Unit"))
        DT_SOTemplate.Columns.Add(New DataColumn("PriceType"))
        DT_SOTemplate.Columns.Add(New DataColumn("BaseUnit", System.Type.GetType("System.Double")))
        DT_SOTemplate.Columns.Add(New DataColumn("VouQty", System.Type.GetType("System.Double")))
        DT_SOTemplate.Columns.Add(New DataColumn("PrimaryQty", System.Type.GetType("System.Double")))
        DT_SOTemplate.Columns.Add(New DataColumn("Price", System.Type.GetType("System.Double")))
        DT_SOTemplate.Columns.Add(New DataColumn("BaseUnitPrice", System.Type.GetType("System.Double")))
        DT_SOTemplate.Columns.Add(New DataColumn("DiscType"))
        DT_SOTemplate.Columns.Add(New DataColumn("DiscPercentage", System.Type.GetType("System.Double")))
        DT_SOTemplate.Columns.Add(New DataColumn("TCDiscountAmount", System.Type.GetType("System.Decimal")))
        DT_SOTemplate.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Decimal")))
        DT_SOTemplate.Columns.Add(New DataColumn("Tax"))
        DT_SOTemplate.Columns.Add(New DataColumn("TaxPercentage", System.Type.GetType("System.Double")))
        DT_SOTemplate.Columns.Add(New DataColumn("TaxAmount", System.Type.GetType("System.Decimal")))
        DT_SOTemplate.Columns.Add(New DataColumn("NonClaimableTaxAmount", System.Type.GetType("System.Decimal")))
        DT_SOTemplate.Columns.Add(New DataColumn("NetAmount", System.Type.GetType("System.Decimal")))
        DT_SOTemplate.Columns.Add(New DataColumn("LCAmount", System.Type.GetType("System.Decimal")))
        DT_SOTemplate.Columns.Add(New DataColumn("LCCostPrice", System.Type.GetType("System.Double")))
        DT_SOTemplate.Columns.Add(New DataColumn("DeliveredTotQty", System.Type.GetType("System.Double")))
        DT_SOTemplate.Columns.Add(New DataColumn("PartNo"))
        DT_SOTemplate.Columns.Add(New DataColumn("Comment"))
        'The below OrgSlno is for quotation to Salesorder
        DT_SOTemplate.Columns.Add(New DataColumn("OrgSlno", System.Type.GetType("System.Int32")))
        DT_SOTemplate.Columns.Add(New DataColumn("Desc1"))
        DT_SOTemplate.Columns.Add(New DataColumn("Desc2"))
        DT_SOTemplate.Columns.Add(New DataColumn("Desc3"))
        DT_SOTemplate.Columns.Add(New DataColumn("Desc4"))
        DT_SOTemplate.Columns.Add(New DataColumn("Desc5"))
        DT_SOTemplate.Columns.Add(New DataColumn("Desc6"))
        DT_SOTemplate.Columns.Add(New DataColumn("Desc7"))
        DT_SOTemplate.Columns.Add(New DataColumn("Desc8"))
        DT_SOTemplate.Columns.Add(New DataColumn("MinSellPrice", System.Type.GetType("System.Decimal")))
        DT_SOTemplate.Columns.Add(New DataColumn("ItemTaxDetails"))
        Return DT_SOTemplate
    End Function
End Class

Public Class csSalesOrderMain
    Public int_BusinessPeriodID As Integer
    Public str_Flag As String
    Public str_FormPrefix As String
    Public str_MenuID As String
    Public str_SalOrd As String
    Public int_RevNo As Integer
    Public dtp_SODate As Date
    Public str_QtnNum As String
    Public int_LedgerID As Integer
    Public str_Alias As String
    Public int_Aging As Integer
    Public str_PayTerm As String
    Public str_Indref As String
    Public str_Comment As String
    Public str_Contact As String
    Public str_SOStatus As String
    Public str_MerchantRef As String
    Public str_SalesManID As String
    Public str_TCCurrency As String
    Public dbl_ExchangeRate As Double
    Public int_StatusCancel As Integer
    Public str_DeliveryAddress As String
    Public str_ContactPerson As String

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

    Public dbl_TCAmount As Double
    Public dbl_TCDisAmount As String
    Public dbl_TCDiscountAmount As Double
    Public dbl_TCAdjAmount As Double
    Public dbl_TCNetAmount As Double
    Public dbl_TCMiscPercentage As String
    Public dbl_TCMiscAmount As Double
    Public dbl_LCNetAmount As Double
    Public dtp_QuotationDate As Date
    Public str_ExpiryDays As String
    Public str_MiscText As String
    Public str_DiscText As String
    Public str_UserComment As String = String.Empty
    Public str_ApproverComment As String = String.Empty
    Public str_ItemTaxCode As String
    Public str_InvoiceTaxCode As String
    Public str_InvoiceTaxXML As String
    Public dbl_TCItemTaxAmount As Double
    Public dbl_TCInvoiceTaxAmount As Double
    Public dt_TaxItemDetails As DataTable
    Public dbl_ItemDiscPercentage As Double
    Public str_WHID As String
    Public str_Consignee As String
    Public str_SalesType As String
    Public str_DeliveryCountry As String
    Public int_LanguageCode As Integer

    Public str_RTF_Description As String
End Class

Public Class csSalesOrderSub
    Public dt_SalesOrderItemDetails
End Class


