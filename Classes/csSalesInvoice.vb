'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class csSalesInvoice
    Inherits csSignature

    Public str_SiteID As String

    Public objCSMain As New csSalesInvoiceMain
    Public objCSSub As New csSalesInvoiceSub
    Public objProject As csProjectDetail

    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        If CustomerSetting.Item("Project").ToString = "True" Then
            objProject = New csProjectDetail
        End If
    End Sub

End Class

Public Class csSalesInvoiceMain
    Public str_Flag As String
    Public str_Prefix As String
    Public int_SeqNo As Integer

    Public str_SISNo As String
    Public int_RevNo As Integer
    Public str_LpoNo As String
    Public str_DONo As String
    Public str_InvRef As String
    Public str_MerchantID As String
    Public str_MerchantName As String
    Public dtp_InvDate As Date
    Public dtp_DueDate As Date
    Public int_Aging As Integer
    Public str_PayTerm As String
    Public str_SalesManID As String
    Public bool_CounterInvoice As Boolean
    Public bool_AffectInventory As Boolean
    Public str_PaymentStatus As String
    Public str_Comment As String

    Public dbl_TCTotalAmount As Double
    Public dbl_TCDisAmount As Double
    Public dbl_TCDiscountAmount As Double
    Public dbl_TCMiscAmount As Double
    Public dbl_TCTaxAmount As Double
    Public dbl_TCNetAmount As Double

    Public dbl_TCPDCAmount As Double

    Public dbl_LCNetAmount As Double
    Public dbl_LCPDCAmount As Double

    Public str_CurrencyID As String
    Public dbl_ExchangeRate As Double

    Public str_CreatedBy As String
    Public dtp_CreatedDate As Date
    Public str_LastUpdatedBy As String
    Public dtp_LastUpdatedDate As Date
    Public str_ApprovedBy As String
    Public dtp_ApprovedDate As Date
    Public bool_ApprovedStatus As Boolean

    Public int_BusinessPeriodID As Integer

    Public str_CashorCredit As String
    Public str_CashLedger As String
End Class

Public Class csSalesInvoiceSub
    Public dt_CS As DataTable
    Public dt_CSMatching As DataTable
End Class






