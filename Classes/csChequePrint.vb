
Public Class csChequePrint
    Inherits csSignature

    Public str_SiteID As String
    Public int_BusinessPeriodID As Integer
    Public str_Flag As String
    Public str_FormPrefix As String

    Public str_VouNo As String
    Public str_ConvertNo As String
    Public int_BankLedgerID As Integer
    Public int_MerchantLedgerID As Integer
    Public str_Alias As String
    Public dtp_VouDate As Date
    Public dtp_ChequeDate As Date
    Public str_ChequeNumber As String
    Public dbl_Amount As Decimal
    Public bool_ACpayeeonly As Boolean
    Public str_Comment As String
    Public int_RevNo As Integer
    Public int_PrintRevNo As Integer
    Public str_RefID As String
    Public bool_Cancelled As Boolean
End Class
