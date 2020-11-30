
Public Class csCustomsInvoice
    Inherits csSignature

    Public str_SiteID As String

    Public objCIMain As New CIMain
    Public DT_CIItemDetails As New DataTable

    Public Function DT_CustomInvoiceTemplate() As DataTable
        DT_CustomInvoiceTemplate = New DataTable

        DT_CustomInvoiceTemplate.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
        DT_CustomInvoiceTemplate.Columns.Add(New DataColumn("InvNo"))
        DT_CustomInvoiceTemplate.Columns.Add(New DataColumn("ItemCode"))
        DT_CustomInvoiceTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_CustomInvoiceTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_CustomInvoiceTemplate.Columns.Add(New DataColumn("Made"))
        DT_CustomInvoiceTemplate.Columns.Add(New DataColumn("Package", System.Type.GetType("System.Int32")))
        DT_CustomInvoiceTemplate.Columns.Add(New DataColumn("Pieces", System.Type.GetType("System.Double")))
        DT_CustomInvoiceTemplate.Columns.Add(New DataColumn("Unit"))
        DT_CustomInvoiceTemplate.Columns.Add(New DataColumn("OrgQty", System.Type.GetType("System.Double")))
        DT_CustomInvoiceTemplate.Columns.Add(New DataColumn("OrgPrice", System.Type.GetType("System.Double")))
        DT_CustomInvoiceTemplate.Columns.Add(New DataColumn("OrgAmount", System.Type.GetType("System.Double")))
        DT_CustomInvoiceTemplate.Columns.Add(New DataColumn("Discount", System.Type.GetType("System.Double")))
        DT_CustomInvoiceTemplate.Columns.Add(New DataColumn("NewQty", System.Type.GetType("System.Double")))
        DT_CustomInvoiceTemplate.Columns.Add(New DataColumn("NewPrice", System.Type.GetType("System.Double")))
        DT_CustomInvoiceTemplate.Columns.Add(New DataColumn("NewAmount", System.Type.GetType("System.Double")))
        Return DT_CustomInvoiceTemplate
    End Function
End Class

Public Class CIMain
    Public str_Flag As String
    Public str_FormPrefix As String
    Public str_MenuID As String
    Public int_BusinessPeriodID As Integer
    Public dtp_Date As Date
    Public int_RevNo As Integer
    Public str_CINo As String
    Public int_LedgerID As Integer
    Public str_Alias As String
    Public str_SalesManID As String
    Public str_Comment As String
    Public str_Cargo As String
    Public dbl_ItemDiscount As Double

    Public str_TCCurrency As String
    Public dbl_ExchangeRate As Double

    Public dbl_TCAmount As Double
    Public dbl_TCDisAmount As Double
    Public dbl_TCDiscountAmount As Double
    Public dbl_TCNetAmount As Double
    Public dbl_LCNetAmount As Double
End Class
