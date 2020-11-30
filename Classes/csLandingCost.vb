
Public Class csLandingCost
    Inherits csSignature

    Public str_SiteID As String
    Public objLCMain As New csLCMain
    Public objLCSub As New csLCSub
   
    Public Function LCSubTemplate() As DataTable
        Dim DT_LCSubTemplate As New DataTable
        DT_LCSubTemplate.Columns.Add(New DataColumn("SlNo", System.Type.GetType("System.Int32")))
        DT_LCSubTemplate.Columns.Add(New DataColumn("LCID"))
        DT_LCSubTemplate.Columns.Add(New DataColumn("VendorLedgerID", System.Type.GetType("System.Int32")))
        DT_LCSubTemplate.Columns.Add(New DataColumn("LedgerType"))
        DT_LCSubTemplate.Columns.Add(New DataColumn("DistType"))
        DT_LCSubTemplate.Columns.Add(New DataColumn("GenExpLedgerID", System.Type.GetType("System.Int32")))
        DT_LCSubTemplate.Columns.Add(New DataColumn("TCAmount", System.Type.GetType("System.Decimal")))
        DT_LCSubTemplate.Columns.Add(New DataColumn("LCAmount", System.Type.GetType("System.Decimal")))
        DT_LCSubTemplate.Columns.Add(New DataColumn("Factor"))
        DT_LCSubTemplate.Columns.Add(New DataColumn("Comment"))
        DT_LCSubTemplate.Columns.Add(New DataColumn("TCPaidAmount", System.Type.GetType("System.Decimal")))
        DT_LCSubTemplate.Columns.Add(New DataColumn("Tax"))
        DT_LCSubTemplate.Columns.Add(New DataColumn("TaxPercentage", System.Type.GetType("System.Decimal")))
        DT_LCSubTemplate.Columns.Add(New DataColumn("TaxAmount", System.Type.GetType("System.Double")))
        DT_LCSubTemplate.Columns.Add(New DataColumn("NonClaimableTaxAmount", System.Type.GetType("System.Decimal")))
        DT_LCSubTemplate.Columns.Add(New DataColumn("ItemTaxDetails"))
        DT_LCSubTemplate.Columns.Add(New DataColumn("TCNetAmount", System.Type.GetType("System.Decimal")))
        DT_LCSubTemplate.Columns.Add(New DataColumn("RefNo"))
        Return DT_LCSubTemplate
    End Function

    Public Function LCItemDetailsTemplate() As DataTable
        Dim dt_LCItemDetailsTemplate As New DataTable

        dt_LCItemDetailsTemplate.Columns.Add(New DataColumn("LCSlNo", System.Type.GetType("System.Int32")))
        dt_LCItemDetailsTemplate.Columns.Add(New DataColumn("ItemCode"))
        dt_LCItemDetailsTemplate.Columns.Add(New DataColumn("Alias1"))
        dt_LCItemDetailsTemplate.Columns.Add(New DataColumn("Alias2"))
        dt_LCItemDetailsTemplate.Columns.Add(New DataColumn("TotalValue", System.Type.GetType("System.Decimal")))
        dt_LCItemDetailsTemplate.Columns.Add(New DataColumn("TotalWeight", System.Type.GetType("System.Decimal")))
        dt_LCItemDetailsTemplate.Columns.Add(New DataColumn("TotalVolume", System.Type.GetType("System.Decimal")))
        dt_LCItemDetailsTemplate.Columns.Add(New DataColumn("UnitLCAddCost", System.Type.GetType("System.Decimal")))
        dt_LCItemDetailsTemplate.Columns.Add(New DataColumn("LCAddCost", System.Type.GetType("System.Decimal")))
        dt_LCItemDetailsTemplate.Columns.Add(New DataColumn("VouNo"))
        dt_LCItemDetailsTemplate.Columns.Add(New DataColumn("VouSlNo", System.Type.GetType("System.Int32")))

        Return dt_LCItemDetailsTemplate
    End Function

    Public Function LCItemSplitValue() As DataTable
        LCItemSplitValue = New DataTable
        LCItemSplitValue.Columns.Add(New DataColumn("LCSlNo", System.Type.GetType("System.Int32")))
        LCItemSplitValue.Columns.Add(New DataColumn("VouSlNo", System.Type.GetType("System.Int32")))
        LCItemSplitValue.Columns.Add(New DataColumn("ItemCode"))
        LCItemSplitValue.Columns.Add(New DataColumn("Alias1"))
        LCItemSplitValue.Columns.Add(New DataColumn("Alias2"))
        LCItemSplitValue.Columns.Add(New DataColumn("SplitColText"))
        LCItemSplitValue.Columns.Add(New DataColumn("SplitColValue", System.Type.GetType("System.Decimal")))
        Return LCItemSplitValue
    End Function

    Public Function LCMrvMainTemplate() As DataTable
        Dim DT_LCMrvDetaisTemplate As New DataTable
        DT_LCMrvDetaisTemplate.Columns.Add(New DataColumn("Vou"))
        DT_LCMrvDetaisTemplate.Columns.Add(New DataColumn("MRV", System.Type.GetType("System.Int32")))
        DT_LCMrvDetaisTemplate.Columns.Add(New DataColumn("MrvDate1", GetType(Date)))
        DT_LCMrvDetaisTemplate.Columns.Add(New DataColumn("Alias"))
        DT_LCMrvDetaisTemplate.Columns.Add(New DataColumn("PIP"))
        DT_LCMrvDetaisTemplate.Columns.Add(New DataColumn("DORef"))
        DT_LCMrvDetaisTemplate.Columns.Add(New DataColumn("TCNetAmount", System.Type.GetType("System.Decimal")))
        Return DT_LCMrvDetaisTemplate
    End Function
    Public Function LCMrvDetaisTemplate() As DataTable
        Dim DT_LCMrvMainTemplate As New DataTable
        DT_LCMrvMainTemplate.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
        DT_LCMrvMainTemplate.Columns.Add(New DataColumn("Vou"))
        DT_LCMrvMainTemplate.Columns.Add(New DataColumn("MRV", System.Type.GetType("System.Int32")))
        DT_LCMrvMainTemplate.Columns.Add(New DataColumn("Alias"))
        DT_LCMrvMainTemplate.Columns.Add(New DataColumn("ItemCode"))
        DT_LCMrvMainTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_LCMrvMainTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_LCMrvMainTemplate.Columns.Add(New DataColumn("Desc1"))
        DT_LCMrvMainTemplate.Columns.Add(New DataColumn("Desc2"))
        DT_LCMrvMainTemplate.Columns.Add(New DataColumn("Desc3"))
        DT_LCMrvMainTemplate.Columns.Add(New DataColumn("Qty", System.Type.GetType("System.Decimal")))
        DT_LCMrvMainTemplate.Columns.Add(New DataColumn("Price", System.Type.GetType("System.Decimal")))
        DT_LCMrvMainTemplate.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Decimal")))
        DT_LCMrvMainTemplate.Columns.Add(New DataColumn("AddCost", System.Type.GetType("System.Decimal")))
        DT_LCMrvMainTemplate.Columns.Add(New DataColumn("AddAmt", System.Type.GetType("System.Decimal")))
        DT_LCMrvMainTemplate.Columns.Add(New DataColumn("FinalPrice", System.Type.GetType("System.Decimal")))
        Return DT_LCMrvMainTemplate
    End Function
End Class

Public Class csLCMain
    Public str_MenuID As String
    Public int_BusinessPeriodID As Integer
    Public str_Flag As String
    Public str_Prefix As String

    Public str_LCNo As String
    Public int_RevNo As Integer
    Public dtp_LCDate As Date
    Public str_VouType As String
    Public str_VouNo As String
    Public str_TCCurrency As String
    Public dbl_ExchangeRate As Double
    Public bool_AffectInventoryCost As Boolean
    Public str_Comment As String
    Public dbl_TCAmount As Double
    Public str_TaxCode As String
    Public dbl_TCTaxAmount As Double
    Public dbl_TCItemTaxAmount As Double
    Public dbl_TCInvoiceTaxAmount As Double
    Public dbl_TCNetAmount As Double
    Public dbl_LCNetAmount As Double

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
    Public str_DiscText As String
    Public bool_TaxFileReturn As Boolean
    Public str_ItemTaxCode As String
    Public dt_TaxItemDetails As DataTable
End Class

Public Class csLCSub
    Public dt_LCSub As DataTable
    Public dt_LCItemDetails As DataTable
    Public dt_LCItemSplitValue As DataTable
End Class
