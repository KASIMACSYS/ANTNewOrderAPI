Public Class csTaxFileAdjustment
    Inherits csSignature
    Public str_CID As String
    Public int_BusinessPeriodID As Integer
    Public ObjTaxFileAdjustment As New csTaxFileAdjustmentMain

End Class
Public Class csTaxFileAdjustmentMain
    Public str_Flag As String
    Public int_TaxID As Integer
    Public str_VouNo As String
    Public str_TaxFileVouNo As String
    Public str_TaxDesc As String
    Public str_TaxLedgerID As String
    Public str_DscLedgerID As String
    Public dbl_AdjustmentAmt As Double
    Public str_Comment As String
    Public str_CreatedBy As String
    Public dtp_VouDate As Date
    Public dtp_CreatedDate As Date = Date.Now
    Public str_LastUpdatedBy As String
    Public dtp_LastUpdatedDate As Date = Date.Now
    Public str_ApprovedBy As String = String.Empty
    Public dtp_ApprovedDate As Date = Date.Now
    Public bool_ApprovedStatus As Integer
    Public dt_TaxFileAdjustment As DataTable
End Class
