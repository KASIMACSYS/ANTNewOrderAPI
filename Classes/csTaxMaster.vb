Public Class csTaxMaster
    Inherits csSignature
    Public str_SiteID As String
    Public int_BusinessPeriodID As Integer
    Public ObjTaxMasterMain As New csTaxMain

End Class

Public Class csTaxMain
    Public str_Flag As String
    Public int_TaxID As Integer
    Public str_TaxCode As String
    Public str_TaxName As String
    Public str_TaxAgent As String
    Public str_TaxDesc As String
    Public str_SalesTax As String
    Public str_PurchaseTax As String
    Public dbl_SalesTaxPercentage As Double
    Public dbl_PurchaseTaxPercentage As Double
    Public bool_PurchaseType As Boolean
    Public bool_InActive As Boolean
    Public dbl_ReverseTax As Double

    Public str_TAN As String
    Public str_TAAN As String
    Public dtp_StartPeriodDate As Date
    Public dtp_EndPeriodDate As Date
    Public dtp_FAFCreationDate As Date
    Public str_FAFVersion As String

    Public str_CreatedBy As String
    Public dtp_CreatedDate As Date = Date.Now
    Public str_LastUpdatedBy As String
    Public dtp_LastUpdatedDate As Date = Date.Now
    Public str_ApprovedBy As String = String.Empty
    Public dtp_ApprovedDate As Date = Date.Now
    Public bool_ApprovedStatus As Integer
    Public dt_TaxMaster As DataTable
End Class
