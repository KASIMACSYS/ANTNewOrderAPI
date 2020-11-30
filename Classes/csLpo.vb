'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Public Class csLpo
    Inherits csSignature
    Public str_CID As String

    Public objLpoMain As New csLpoMain
    Public objLpoSub As New csLpoSub
    Public objproject As New csProjectDetail

    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        ''If CustomerSetting.Item("Project").ToString = "True" Then
        objproject = New csProjectDetail
        ''End If
    End Sub


    Public Function DBTemplate() As DataTable
        Dim DT_LPOTemplate As New DataTable
        DT_LPOTemplate.Columns.Add(New DataColumn("SortNo", System.Type.GetType("System.Int32")))
        DT_LPOTemplate.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
        DT_LPOTemplate.Columns.Add(New DataColumn("BarCodeNo"))
        DT_LPOTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_LPOTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_LPOTemplate.Columns.Add(New DataColumn("ItemCode"))
        DT_LPOTemplate.Columns.Add(New DataColumn("Unit"))
        DT_LPOTemplate.Columns.Add(New DataColumn("BaseUnit", System.Type.GetType("System.Double")))
        DT_LPOTemplate.Columns.Add(New DataColumn("Package", System.Type.GetType("System.Double")))
        DT_LPOTemplate.Columns.Add(New DataColumn("Pieces", System.Type.GetType("System.Double")))
        DT_LPOTemplate.Columns.Add(New DataColumn("VouQty", System.Type.GetType("System.Double")))
        DT_LPOTemplate.Columns.Add(New DataColumn("PrimaryQty", System.Type.GetType("System.Double")))
        DT_LPOTemplate.Columns.Add(New DataColumn("PriceType"))
        DT_LPOTemplate.Columns.Add(New DataColumn("Price", System.Type.GetType("System.Double")))
        DT_LPOTemplate.Columns.Add(New DataColumn("BaseUnitPrice", System.Type.GetType("System.Double")))
        DT_LPOTemplate.Columns.Add(New DataColumn("DiscType"))
        DT_LPOTemplate.Columns.Add(New DataColumn("DiscPercentage", System.Type.GetType("System.Double")))
        DT_LPOTemplate.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Decimal")))
        DT_LPOTemplate.Columns.Add(New DataColumn("Tax"))
        DT_LPOTemplate.Columns.Add(New DataColumn("TaxPercentage", System.Type.GetType("System.Double")))
        DT_LPOTemplate.Columns.Add(New DataColumn("TaxAmount", System.Type.GetType("System.Decimal")))
        DT_LPOTemplate.Columns.Add(New DataColumn("NonClaimableTaxAmount", System.Type.GetType("System.Decimal")))
        DT_LPOTemplate.Columns.Add(New DataColumn("NetAmount", System.Type.GetType("System.Decimal")))
        DT_LPOTemplate.Columns.Add(New DataColumn("TCDiscountAmount", System.Type.GetType("System.Decimal")))
        DT_LPOTemplate.Columns.Add(New DataColumn("LCAmount", System.Type.GetType("System.Decimal")))
        DT_LPOTemplate.Columns.Add(New DataColumn("Comment"))
        DT_LPOTemplate.Columns.Add(New DataColumn("PartNo"))
        'The below OrgSlno is for Indent to Lpo

        DT_LPOTemplate.Columns.Add(New DataColumn("OrgSlno", System.Type.GetType("System.Int32")))
        DT_LPOTemplate.Columns.Add(New DataColumn("DeliveredTotQty", System.Type.GetType("System.Double")))
        DT_LPOTemplate.Columns.Add(New DataColumn("Desc1"))
        DT_LPOTemplate.Columns.Add(New DataColumn("Desc2"))
        DT_LPOTemplate.Columns.Add(New DataColumn("Desc3"))
        DT_LPOTemplate.Columns.Add(New DataColumn("Desc4"))
        DT_LPOTemplate.Columns.Add(New DataColumn("Desc5"))
        DT_LPOTemplate.Columns.Add(New DataColumn("Desc6"))
        DT_LPOTemplate.Columns.Add(New DataColumn("Desc7"))
        DT_LPOTemplate.Columns.Add(New DataColumn("Desc8"))
        DT_LPOTemplate.Columns.Add(New DataColumn("ItemTaxDetails"))
        DT_LPOTemplate.Columns.Add(New DataColumn("ConvertNo"))
        Return DT_LPOTemplate
    End Function
End Class

Public Class csLpoMain
    Public str_MenuID As String
    Public str_ConvertFrom As String
    Public str_Flag As String
    Public str_FormPrefix As String
    Public int_BusinessPeriodID As Integer
    Public str_LpoNo As String
    Public dtp_LpoDate1 As Date
    Public dtp_LpoDate2 As Date
    Public str_IndentNo As String
    Public str_EnqNo As String
    Public int_LedgerID As Integer
    Public str_Alias As String
    Public int_Aging As Integer
    Public str_PayTerm As String
    Public str_IndRef As String
    Public str_DelivAddress As String
    Public str_Comment As String
    Public str_RefNo As String
    Public str_Contact As String
    Public int_RevNo As Integer
    Public str_LpoStatus As String
    Public int_StatusCancel As Integer
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
    Public dbl_TCMiscPercentage As String
    Public dbl_TCMiscAmount As Double
    Public dbl_TCDiscountAmount As Double
    Public dbl_TCItemTaxAmount As Double
    Public dbl_TCInvoiceTaxAmount As Double
    Public dbl_TCAdjAmount As Double
    Public dbl_TCNetAmount As Double
    Public dbl_LCNetAmount As Double
    Public str_TCCurrency As String
    Public dbl_ExchangeRate As Double
    Public dtp_IndentDate As Date
    Public str_ExpiryDays As String
    Public str_MiscText As String
    Public str_DiscText As String
    Public str_UserComment As String = String.Empty
    Public str_TaxCode As String
    Public dbl_ItemDiscPercentage As Double
    Public str_ItemTaxCode As String
    Public str_InvoiceTaxCode As String
    Public str_PermitNo As String
    Public str_InvoiceTaxXML As String
    Public dt_TaxItemDetails As DataTable
    Public int_LanguageCode As Integer
End Class

Public Class csLpoSub
    Public dt_Lpo As DataTable
End Class





