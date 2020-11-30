'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Public Class csQuotation
    Inherits csSignature

    Public str_CID As String
    Public objQuotationMain As New csQuotationMain
    Public objQuotationSub As New csQuotationSub
    Public objproject As New csProjectDetail
    Public DTItemExtraDetails As DataTable

    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        'If CustomerSetting.Item("Project").ToString = "True" Then
        objproject = New csProjectDetail
        ' End If
    End Sub

    Public Function DBTemplate() As DataTable
        Dim DT_QTNTemplate As New DataTable
        DT_QTNTemplate.Columns.Add(New DataColumn("SortNo", System.Type.GetType("System.Int32")))
        DT_QTNTemplate.Columns.Add(New DataColumn("SlNo", System.Type.GetType("System.Int32")))
        DT_QTNTemplate.Columns.Add(New DataColumn("BarCodeNo"))
        DT_QTNTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_QTNTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_QTNTemplate.Columns.Add(New DataColumn("ItemCode"))
        DT_QTNTemplate.Columns.Add(New DataColumn("Package", System.Type.GetType("System.Double")))
        DT_QTNTemplate.Columns.Add(New DataColumn("Pieces", System.Type.GetType("System.Double")))
        DT_QTNTemplate.Columns.Add(New DataColumn("Unit"))
        DT_QTNTemplate.Columns.Add(New DataColumn("PriceType"))
        DT_QTNTemplate.Columns.Add(New DataColumn("BaseUnit", System.Type.GetType("System.Double")))
        DT_QTNTemplate.Columns.Add(New DataColumn("VouQty", System.Type.GetType("System.Double")))
        DT_QTNTemplate.Columns.Add(New DataColumn("PrimaryQty", System.Type.GetType("System.Double")))
        DT_QTNTemplate.Columns.Add(New DataColumn("Price", System.Type.GetType("System.Double")))
        DT_QTNTemplate.Columns.Add(New DataColumn("BaseUnitPrice", System.Type.GetType("System.Double")))
        DT_QTNTemplate.Columns.Add(New DataColumn("DiscType"))
        DT_QTNTemplate.Columns.Add(New DataColumn("DiscPercentage", System.Type.GetType("System.Decimal")))
        DT_QTNTemplate.Columns.Add(New DataColumn("TCDiscountAmount", System.Type.GetType("System.Decimal")))
        DT_QTNTemplate.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Decimal")))
        DT_QTNTemplate.Columns.Add(New DataColumn("Tax"))
        DT_QTNTemplate.Columns.Add(New DataColumn("TaxPercentage", System.Type.GetType("System.Double")))
        DT_QTNTemplate.Columns.Add(New DataColumn("TaxAmount", System.Type.GetType("System.Decimal")))
        DT_QTNTemplate.Columns.Add(New DataColumn("NonClaimableTaxAmount", System.Type.GetType("System.Decimal")))
        DT_QTNTemplate.Columns.Add(New DataColumn("NetAmount", System.Type.GetType("System.Decimal")))
        DT_QTNTemplate.Columns.Add(New DataColumn("LCAmount", System.Type.GetType("System.Decimal")))
        DT_QTNTemplate.Columns.Add(New DataColumn("LCCostPrice", System.Type.GetType("System.Double")))
        DT_QTNTemplate.Columns.Add(New DataColumn("DeliveredTotQty", System.Type.GetType("System.Double")))
        DT_QTNTemplate.Columns.Add(New DataColumn("PartNo"))
        DT_QTNTemplate.Columns.Add(New DataColumn("Comment"))
        DT_QTNTemplate.Columns.Add(New DataColumn("Desc1"))
        DT_QTNTemplate.Columns.Add(New DataColumn("Desc2"))
        DT_QTNTemplate.Columns.Add(New DataColumn("Desc3"))
        DT_QTNTemplate.Columns.Add(New DataColumn("Desc4"))
        DT_QTNTemplate.Columns.Add(New DataColumn("Desc5"))
        DT_QTNTemplate.Columns.Add(New DataColumn("Desc6"))
        DT_QTNTemplate.Columns.Add(New DataColumn("Desc7"))
        DT_QTNTemplate.Columns.Add(New DataColumn("Desc8"))
        DT_QTNTemplate.Columns.Add(New DataColumn("MinSellPrice", System.Type.GetType("System.Decimal")))
        DT_QTNTemplate.Columns.Add(New DataColumn("ItemTaxDetails"))
        Return DT_QTNTemplate
    End Function
End Class

Public Class csQuotationMain
    Public int_BusinessPeriodID As Integer
    Public str_Flag As String
    Public str_MenuID As String
    Public str_FormPrefix As String

    Public str_QtnNo As String
    Public int_RevNo As Integer
    Public dtp_QtnDate As Date
    Public int_LedgerID As Integer
    Public str_Alias As String
    Public int_Aging As Integer
    Public str_PayTerm As String

    'Public str_Surface As String
    Public str_IndRef As String
    Public str_DeliverIn As String
    Public str_QtnValidity As String
    Public Str_QtnStatus As String
    Public str_Comment As String
    Public str_Contact As String
    Public str_SalesManID As String
    Public str_TCCurrency As String
    Public dbl_ExchangeRate As Double
    Public int_StatusCancel As Integer

    Public dbl_TCAmount As Double
    Public dbl_TCDisAmount As String
    Public dbl_TCDiscountAmount As Double
    Public dbl_TCAdjAmount As Double
    Public dbl_TCNetAmount As Double
    Public dbl_TCMiscPercentage As String
    Public dbl_TCMiscAmount As Double
    Public dbl_LCNetAmount As Double
    Public str_MiscText As String
    Public str_DiscText As String
    Public str_Desc1 As String
    Public str_Desc2 As String
    Public str_Desc3 As String
    Public str_Desc4 As String
    Public str_Desc5 As String
    Public str_Desc6 As String
    Public str_Desc7 As String
    Public str_Desc8 As String
    Public str_ExpiryDays As String
    Public str_EstNo As String
    Public str_ItemTaxCode As String
    Public str_InvoiceTaxCode As String
    Public str_InvoiceTaxXML As String
    Public dbl_TCItemTaxAmount As Double
    Public dbl_TCInvoiceTaxAmount As Double
    Public dt_TaxItemDetails As DataTable
    Public dbl_ItemDiscPercentage As Double
    Public int_LanguageCode As Integer
    Public int_RevisionHistoryNo As Integer

    Public _XMLCustomData As String

    Public str_UserComment As String = String.Empty
    Public str_ApproverComment As String = String.Empty

    Public str_RTF_Description As String
End Class

Public Class csQuotationSub
    Public dt_Quotation As DataTable
End Class






