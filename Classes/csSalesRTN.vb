'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class csSalesRTN
    Inherits csSignature
    Public str_SiteID As String

    Public objSalRtnMain As New csSalRtnMain
    Public objSalRtnSub As New csSalRtnSub
    Public objproject As New csProjectDetail
    Public DTBatch As New DataTable

    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        '' If CustomerSetting.Item("Project").ToString = "True" Then
        objproject = New csProjectDetail
        '' End If
    End Sub
    Public Function DBTemplate() As DataTable
        Dim DT_SRTTemplate As New DataTable
        DT_SRTTemplate.Columns.Add(New DataColumn("SortNo", System.Type.GetType("System.Int32")))
        DT_SRTTemplate.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
        DT_SRTTemplate.Columns("Slno").AutoIncrement = True
        DT_SRTTemplate.Columns("Slno").AutoIncrementSeed = 1
        DT_SRTTemplate.Columns("Slno").AutoIncrementStep = 1
        DT_SRTTemplate.Columns.Add(New DataColumn("BarCodeNo"))
        DT_SRTTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_SRTTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_SRTTemplate.Columns.Add(New DataColumn("ItemCode"))
        DT_SRTTemplate.Columns.Add(New DataColumn("Unit"))
        DT_SRTTemplate.Columns.Add(New DataColumn("BaseUnit", System.Type.GetType("System.Double")))
        DT_SRTTemplate.Columns.Add(New DataColumn("VouQty", System.Type.GetType("System.Double")))
        DT_SRTTemplate.Columns.Add(New DataColumn("PrimaryQty", System.Type.GetType("System.Double")))
        DT_SRTTemplate.Columns.Add(New DataColumn("Price", System.Type.GetType("System.Double")))
        DT_SRTTemplate.Columns.Add(New DataColumn("BaseUnitPrice", System.Type.GetType("System.Decimal")))
        DT_SRTTemplate.Columns.Add(New DataColumn("DiscType"))
        DT_SRTTemplate.Columns.Add(New DataColumn("DiscPercentage", System.Type.GetType("System.Double")))
        DT_SRTTemplate.Columns.Add(New DataColumn("TCDiscountAmount", System.Type.GetType("System.Decimal")))
        DT_SRTTemplate.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Decimal")))
        DT_SRTTemplate.Columns.Add(New DataColumn("Tax"))
        DT_SRTTemplate.Columns.Add(New DataColumn("TaxPercentage", System.Type.GetType("System.Double")))
        DT_SRTTemplate.Columns.Add(New DataColumn("TaxAmount", System.Type.GetType("System.Decimal")))
        DT_SRTTemplate.Columns.Add(New DataColumn("NonClaimableTaxAmount", System.Type.GetType("System.Decimal")))
        DT_SRTTemplate.Columns.Add(New DataColumn("NetAmount", System.Type.GetType("System.Decimal")))
        DT_SRTTemplate.Columns.Add(New DataColumn("LCCostPrice", System.Type.GetType("System.Decimal")))
        DT_SRTTemplate.Columns.Add(New DataColumn("LCCostAmount", System.Type.GetType("System.Decimal")))
        DT_SRTTemplate.Columns.Add(New DataColumn("LCAmount", System.Type.GetType("System.Decimal")))
        DT_SRTTemplate.Columns.Add(New DataColumn("Desc1"))
        DT_SRTTemplate.Columns.Add(New DataColumn("Desc2"))
        DT_SRTTemplate.Columns.Add(New DataColumn("Desc3"))
        DT_SRTTemplate.Columns.Add(New DataColumn("Desc4"))
        DT_SRTTemplate.Columns.Add(New DataColumn("Desc5"))
        DT_SRTTemplate.Columns.Add(New DataColumn("Desc6"))
        DT_SRTTemplate.Columns.Add(New DataColumn("Desc7"))
        DT_SRTTemplate.Columns.Add(New DataColumn("Desc8"))
        DT_SRTTemplate.Columns.Add(New DataColumn("PartNo"))
        DT_SRTTemplate.Columns.Add(New DataColumn("SerialNo"))
        DT_SRTTemplate.Columns.Add(New DataColumn("Comment"))
        DT_SRTTemplate.Columns.Add(New DataColumn("ItemTaxDetails"))
        Return DT_SRTTemplate
    End Function
End Class

Public Class csSalRtnMain
    Public str_MenuID As String
    Public int_BusinessPeriodID As Integer
    Public str_SRNo As String
    Public int_RevNo As Integer
    Public dtp_RTNDate1 As Date
    Public dtp_RTNDate2 As Date
    Public str_InvRef As String
    Public str_DoRef As String
    Public str_LpoRef As String
    Public int_LedgerID As Integer
    Public str_Alias As String
    Public str_SalesManID As String
    Public int_Aging As Integer
    Public str_PayTerm As String
    Public str_Ref As String
    Public str_Comment As String
    Public bool_RtnStatus As Boolean
    Public dbl_TCAmount As Double
    Public dbl_TCDisAmount As Double
    Public dbl_TCDiscountAmount As Double
    Public dbl_TCAdjAMount As Double
    Public dbl_TCNetAmount As Double
    Public dbl_LCNetAmount As Double
    Public str_TCCurrency As String
    Public dbl_ExchangeRate As Double
    Public int_StatusCancel As Integer
    Public int_LanguageCode As Integer

    Public str_CurrencyCode As String
    Public bool_ApprovedStatus As Boolean
    Public str_Flag As String
    Public str_SalesRtnPrefix As String
    Public str_CurrencyID As String

    Public dt_InvoiceAccounts As DataTable
    Public str_WHID As String
    Public str_RtnVouType As String
    Public str_RtnVouTypeNo As String
    Public str_DiscText As String
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
    Public str_UserComment As String = String.Empty
    Public bool_TaxFileReturn As Boolean
    Public str_ItemTaxCode As String
    Public str_InvoiceTaxCode As String
    Public str_InvoiceTaxXML As String
    Public dbl_TCItemTaxAmount As Double
    Public dbl_TCInvoiceTaxAmount As Double
    Public dt_TaxItemDetails As DataTable
    Public dt_InvoiceAccountsCostCentre As New DataTable
End Class

Public Class csSalRtnSub
    Public dt_SalRtn As DataTable
    Public dt_SRTMatching As DataTable
End Class


