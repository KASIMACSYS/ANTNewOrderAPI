'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class csPurchaseRTN
    Inherits csSignature
    Public str_SiteID As String

    Public objPurRtnMain As New csPurRtnMain
    Public objPurRtnSub As New csPurRtnSub
    Public objproject As New csProjectDetail
    Public DTBatch As New DataTable

    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        ''If CustomerSetting.Item("Project").ToString = "True" Then
        objproject = New csProjectDetail
        '' End If
    End Sub

    Public Function DBTemplate() As DataTable
        Dim DT_PRTTemplate As New DataTable
        DT_PRTTemplate.Columns.Add(New DataColumn("SortNo", System.Type.GetType("System.Int32")))
        DT_PRTTemplate.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
        DT_PRTTemplate.Columns("Slno").AutoIncrement = True
        DT_PRTTemplate.Columns("Slno").AutoIncrementSeed = 1
        DT_PRTTemplate.Columns("Slno").AutoIncrementStep = 1
        DT_PRTTemplate.Columns.Add(New DataColumn("BarCodeNo"))
        DT_PRTTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_PRTTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_PRTTemplate.Columns.Add(New DataColumn("ItemCode"))
        DT_PRTTemplate.Columns.Add(New DataColumn("Unit"))
        DT_PRTTemplate.Columns.Add(New DataColumn("BaseUnit", System.Type.GetType("System.Double")))
        DT_PRTTemplate.Columns.Add(New DataColumn("VouQty", System.Type.GetType("System.Double")))
        DT_PRTTemplate.Columns.Add(New DataColumn("PrimaryQty", System.Type.GetType("System.Double")))
        DT_PRTTemplate.Columns.Add(New DataColumn("Price", System.Type.GetType("System.Double")))
        DT_PRTTemplate.Columns.Add(New DataColumn("BaseUnitPrice", System.Type.GetType("System.Double")))
        DT_PRTTemplate.Columns.Add(New DataColumn("DiscType"))
        DT_PRTTemplate.Columns.Add(New DataColumn("DiscPercentage", System.Type.GetType("System.Double")))
        DT_PRTTemplate.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Double")))
        DT_PRTTemplate.Columns.Add(New DataColumn("Tax"))
        DT_PRTTemplate.Columns.Add(New DataColumn("TaxPercentage", System.Type.GetType("System.Double")))
        DT_PRTTemplate.Columns.Add(New DataColumn("TaxAmount", System.Type.GetType("System.Decimal")))
        DT_PRTTemplate.Columns.Add(New DataColumn("NonClaimableTaxAmount", System.Type.GetType("System.Decimal")))
        DT_PRTTemplate.Columns.Add(New DataColumn("NetAmount", System.Type.GetType("System.Decimal")))
        DT_PRTTemplate.Columns.Add(New DataColumn("LCAmount", System.Type.GetType("System.Double")))
        DT_PRTTemplate.Columns.Add(New DataColumn("TCDiscountAmount", System.Type.GetType("System.Double")))
        DT_PRTTemplate.Columns.Add(New DataColumn("Desc1"))
        DT_PRTTemplate.Columns.Add(New DataColumn("Desc2"))
        DT_PRTTemplate.Columns.Add(New DataColumn("Desc3"))
        DT_PRTTemplate.Columns.Add(New DataColumn("Desc4"))
        DT_PRTTemplate.Columns.Add(New DataColumn("Desc5"))
        DT_PRTTemplate.Columns.Add(New DataColumn("Desc6"))
        DT_PRTTemplate.Columns.Add(New DataColumn("Desc7"))
        DT_PRTTemplate.Columns.Add(New DataColumn("Desc8"))
        DT_PRTTemplate.Columns.Add(New DataColumn("PartNo"))
        DT_PRTTemplate.Columns.Add(New DataColumn("SerialNo"))
        DT_PRTTemplate.Columns.Add(New DataColumn("Comment"))
        DT_PRTTemplate.Columns.Add(New DataColumn("ItemTaxDetails"))
        Return DT_PRTTemplate
    End Function
End Class

Public Class csPurRtnMain
    Public str_MenuID As String
    Public int_BusinessPeriodID As Integer
    Public str_PRNo As String
    Public int_RevNo As Integer
    Public str_InvRef As String
    Public str_MrvNo As String
    Public str_LpoNo As String
    Public dtp_RtnDate1 As Date
    Public dtp_RtnDate2 As Date
    Public int_LedgerID As Integer
    Public str_Alias As String
    Public int_Aging As Integer
    Public str_PayTerm As String
    Public str_Comment As String
    Public bool_RtnStatus As Boolean
    Public str_TCCurrency As String
    Public dbl_ExchangeRate As Double
    Public dbl_TCAmount As Double
    Public dbl_TCDisAmount As Double
    Public dbl_TCDiscount As Double
    Public dbl_TCNetAmount As Double
    Public dbl_TCBalAmount As Double
    Public dbl_TCAdjAmount As Double
    Public dbl_LCNetAmount As Double
    Public int_StatusCancel As Integer
    Public str_RtnVouType As String
    Public str_RtnVouTypeNo As String
    Public bool_ApprovedStatus As Boolean
    Public str_Flag As String
    Public str_PurchaseRTNPrefix As String
    Public dt_InvoiceAccounts As DataTable
    Public int_LanguageCode As Integer

    Public str_WHID As String
    Public str_Desc1 As String
    Public str_Desc2 As String
    Public str_Desc3 As String
    Public str_Desc4 As String
    Public str_Desc5 As String
    Public str_Desc6 As String
    Public str_Desc7 As String
    Public str_Desc8 As String
    Public str_UserComment As String = String.Empty
    Public bool_TaxFileReturn As Boolean
    Public str_ItemTaxCode As String
    Public str_InvoiceTaxCode As String
    Public str_InvoiceTaxXML As String
    Public dbl_TCItemTaxAmount As Double
    Public dbl_TCInvoiceTaxAmount As Double
    Public str_DiscText As String

    Public dt_TaxItemDetails As DataTable
End Class

Public Class csPurRtnSub
    Public dt_PurRtn As DataTable
    Public dt_PRTMatching As DataTable
End Class



