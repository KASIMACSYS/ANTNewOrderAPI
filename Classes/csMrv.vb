'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class csMrv
    Inherits csSignature

    Public int_CID As String

    Public objMrvMain As New csMrvMain
    Public objMrvSub As New csMrvSub
    Public objproject As New csProjectDetail
    Public DTBatch As DataTable

    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        '' If CustomerSetting.Item("Project").ToString = "True" Then
        objproject = New csProjectDetail
        '' End If
    End Sub

    Public Function DBTemplate() As DataTable
        Dim DT_MRVTemplate As New DataTable
        DT_MRVTemplate.Columns.Add(New DataColumn("SortNo", System.Type.GetType("System.Int32")))
        DT_MRVTemplate.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
        DT_MRVTemplate.Columns.Add(New DataColumn("BarCodeNo"))
        DT_MRVTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_MRVTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_MRVTemplate.Columns.Add(New DataColumn("ItemCode"))
        DT_MRVTemplate.Columns.Add(New DataColumn("Package", System.Type.GetType("System.Double")))
        DT_MRVTemplate.Columns.Add(New DataColumn("Pieces", System.Type.GetType("System.Double")))
        DT_MRVTemplate.Columns.Add(New DataColumn("Unit"))
        DT_MRVTemplate.Columns.Add(New DataColumn("BaseUnit", System.Type.GetType("System.Double")))
        DT_MRVTemplate.Columns.Add(New DataColumn("VouQty", System.Type.GetType("System.Double")))
        DT_MRVTemplate.Columns.Add(New DataColumn("PrimaryQty", System.Type.GetType("System.Double")))
        DT_MRVTemplate.Columns.Add(New DataColumn("PriceType"))
        DT_MRVTemplate.Columns.Add(New DataColumn("Price", System.Type.GetType("System.Double")))
        DT_MRVTemplate.Columns.Add(New DataColumn("BaseUnitPrice", System.Type.GetType("System.Double")))
        DT_MRVTemplate.Columns.Add(New DataColumn("DiscType"))
        DT_MRVTemplate.Columns.Add(New DataColumn("DiscPercentage", System.Type.GetType("System.Double")))
        DT_MRVTemplate.Columns.Add(New DataColumn("TCDiscountAmount", System.Type.GetType("System.Decimal")))
        DT_MRVTemplate.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Decimal")))
        DT_MRVTemplate.Columns.Add(New DataColumn("Tax"))
        DT_MRVTemplate.Columns.Add(New DataColumn("TaxPercentage", System.Type.GetType("System.Double")))
        DT_MRVTemplate.Columns.Add(New DataColumn("TaxAmount", System.Type.GetType("System.Decimal")))
        DT_MRVTemplate.Columns.Add(New DataColumn("NonClaimableTaxAmount", System.Type.GetType("System.Boolean")))
        'DT_MRVTemplate.Columns.Add(New DataColumn("Tax2Code"))
        'DT_MRVTemplate.Columns.Add(New DataColumn("Tax2Percentage", System.Type.GetType("System.Double")))
        'DT_MRVTemplate.Columns.Add(New DataColumn("Tax2Amount", System.Type.GetType("System.Decimal")))
        'DT_MRVTemplate.Columns.Add(New DataColumn("IncludeTax2InCost", System.Type.GetType("System.Boolean")))
        DT_MRVTemplate.Columns.Add(New DataColumn("NetAmount", System.Type.GetType("System.Decimal")))
        DT_MRVTemplate.Columns.Add(New DataColumn("LCAmount", System.Type.GetType("System.Decimal")))
        DT_MRVTemplate.Columns.Add(New DataColumn("LCCostPrice", System.Type.GetType("System.Double")))
  		DT_MRVTemplate.Columns.Add(New DataColumn("TCDiscAmt", System.Type.GetType("System.Double")))
        DT_MRVTemplate.Columns.Add(New DataColumn("LCDiscAmt", System.Type.GetType("System.Double")))
        DT_MRVTemplate.Columns.Add(New DataColumn("LCAddCost", System.Type.GetType("System.Double")))
     
        DT_MRVTemplate.Columns.Add(New DataColumn("Comment"))
        DT_MRVTemplate.Columns.Add(New DataColumn("PartNo"))
        DT_MRVTemplate.Columns.Add(New DataColumn("POQty", System.Type.GetType("System.Double")))
        DT_MRVTemplate.Columns.Add(New DataColumn("BalanceQty", System.Type.GetType("System.Double")))
        'DT_MRVTemplate.Columns.Add(New DataColumn("Tax", System.Type.GetType("System.Double")))
        DT_MRVTemplate.Columns.Add(New DataColumn("WHID"))
        DT_MRVTemplate.Columns.Add(New DataColumn("OrgSlno", System.Type.GetType("System.Int32")))
        DT_MRVTemplate.Columns.Add(New DataColumn("Desc1"))
        DT_MRVTemplate.Columns.Add(New DataColumn("Desc2"))
        DT_MRVTemplate.Columns.Add(New DataColumn("Desc3"))
        DT_MRVTemplate.Columns.Add(New DataColumn("Desc4"))
        DT_MRVTemplate.Columns.Add(New DataColumn("Desc5"))
        DT_MRVTemplate.Columns.Add(New DataColumn("Desc6"))
        DT_MRVTemplate.Columns.Add(New DataColumn("Desc7"))
        DT_MRVTemplate.Columns.Add(New DataColumn("Desc8"))
        'DT_MRVTemplate.Columns.Add(New DataColumn("SerialNo", System.Type.GetType("System.Int32")))
        DT_MRVTemplate.Columns.Add(New DataColumn("ItemTaxDetails"))
        DT_MRVTemplate.Columns.Add(New DataColumn("ConvertNo"))
        Return DT_MRVTemplate
    End Function
End Class

Public Class csMrvMain
    Public str_Flag As String
    Public str_MenuID As String
    Public str_FormPrefix As String
    Public int_BusinessPeriodID As Integer
    Public str_MrvNo As String
    Public int_RevNo As Integer
    Public str_LpoNo As String
    Public str_DoNo As String
    Public dtp_MrvDate1 As Date
    Public dtp_MrvDate2 As Date
    Public int_LedgerID As Integer
    Public str_Alias As String
    Public int_Aging As Integer
    Public str_PayTerm As String
    Public str_Comment As String
    Public bool_StatusInvMatched As Boolean
    Public int_StatusCancel As Integer
    Public str_Pin As String
    Public str_PIP As String
    Public str_PayCertComment As String
    Public bool_ConvertLpo As Boolean
    Public bool_ConvertInv As Boolean
    Public str_TCCurrency As String
    Public dbl_ExchangeRate As Double

    Public str_Desc1 As String
    Public str_Desc2 As String
    Public str_Desc3 As String
    Public str_Desc4 As String
    Public str_Desc5 As String
    Public str_Desc6 As String
    Public str_Desc7 As String
    Public str_Desc8 As String

    Public dbl_TCAmount As Double
    Public dbl_TCDisAmount As String
    Public dbl_TCDiscountAmount As Double
    Public dbl_TCMiscPercentage As String
    Public dbl_TCMiscAmount As Double
    Public dbl_TCNetAmount As Double
    Public dbl_LCNetAmount As Double
    Public dbl_TCVatAmount As Double
    Public dbl_TCItemTaxAmount As Double
    Public dbl_TCInvoiceTaxAmount As Double
    Public dbl_TCAdjAmount As Double
    Public dbl_LCLandingCost As Double
    Public dbl_PIPAmount As Double
    Public str_MiscText As String
    Public str_DiscText As String

    Public str_ConvertForm As String
    Public bool_Status As Boolean ' TODO
    Public bool_RStatus As Boolean ' TODO
    Public str_UserComment As String = String.Empty
    Public str_WHID As String
    Public dbl_ItemDiscPercentage As Double

    Public str_ItemTaxCode As String
    Public str_InvoiceTaxCode As String
    Public str_InvoiceTaxXML As String
    Public int_LanguageCode As Integer
    Public str_DeliveryAddress As String
    Public str_ContactPerson As String
    Public dt_TaxItemDetails As DataTable

End Class

Public Class csMrvSub
    Public dt_MrvSub As DataTable
    Public dt_EditMrvSub As DataTable
End Class




