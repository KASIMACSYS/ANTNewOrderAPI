
Public Class csEnquiryForm
    Inherits csSignature
    Public str_SiteID As String

    Public objEnquiryMain As New csEnquiryMain
    Public objEnquirySub As New csEnquirySub
    Public objproject As New csProjectDetail

    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        ''If CustomerSetting.Item("Project").ToString = "True" Then
        objproject = New csProjectDetail
        ''End If
    End Sub


    Public Function DBTemplate() As DataTable
        Dim DT_EnquiryTemplate As New DataTable
        DT_EnquiryTemplate.Columns.Add(New DataColumn("SortNo", System.Type.GetType("System.Int32")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("BarCodeNo"))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("ItemDesc"))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("ItemCode"))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("Unit"))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("BaseUnit", System.Type.GetType("System.Double")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("Package", System.Type.GetType("System.Double")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("Pieces", System.Type.GetType("System.Double")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("VouQty", System.Type.GetType("System.Double")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("PrimaryQty", System.Type.GetType("System.Double")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("PriceType"))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("Price", System.Type.GetType("System.Double")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("BaseUnitPrice", System.Type.GetType("System.Double")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("DiscType"))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("DiscPercentage", System.Type.GetType("System.Double")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Double")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("TaxPercentage", System.Type.GetType("System.Double")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("TaxAmount", System.Type.GetType("System.Decimal")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("NetAmount", System.Type.GetType("System.Decimal")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("TCDiscountAmount", System.Type.GetType("System.Double")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("LCAmount", System.Type.GetType("System.Double")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("Comment"))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("PartNo"))
        'The below OrgSlno is for Indent to Enquiry

        DT_EnquiryTemplate.Columns.Add(New DataColumn("OrgSlno", System.Type.GetType("System.Int32")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("DeliveredTotQty", System.Type.GetType("System.Double")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("VenQty", System.Type.GetType("System.Double")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("OrderQty", System.Type.GetType("System.Double")))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("VenItemCode"))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("Desc1"))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("Desc2"))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("Desc3"))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("Desc4"))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("Desc5"))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("Desc6"))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("Desc7"))
        DT_EnquiryTemplate.Columns.Add(New DataColumn("Desc8"))

        Return DT_EnquiryTemplate
    End Function
End Class

Public Class csEnquiryMain
    Public str_Flag As String
    Public str_FormPrefix As String
    Public int_BusinessPeriodID As Integer
    Public str_EnquiryNo As String
    Public dtp_EnquiryDate1 As Date
    Public dtp_EnquiryDate2 As Date
    Public str_IndentNo As String
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
    Public str_EnquiryStatus As String
    Public bool_StatusCancel As Boolean
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
    Public dbl_TCNetAmount As Double
    Public dbl_LCNetAmount As Double
    Public str_TCCurrency As String
    Public dbl_ExchangeRate As Double
    Public dtp_IndentDate As Date
    Public str_ExpiryDays As String
    Public str_MiscText As String
End Class

Public Class csEnquirySub
    Public dt_Enquiry As DataTable
End Class
