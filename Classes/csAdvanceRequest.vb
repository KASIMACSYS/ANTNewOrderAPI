Public Class csAdvanceRequest
    Inherits csSignature
    Public str_CID As String
    Public int_BusinessPeriodID As Integer
    Public ObjAdvanceRequestMain As New csAdvanceRequestMain
    Public ObjAdvanceRequestSub As New csAdvanceRequestSub
    Public Class csAdvanceRequestMain
        Public Str_Flag As String
        Public Str_MenuID As String
        Public Str_FormPrefix As String
        Public Str_ReqID As String
        Public Str_PaymentReqID As String
        Public int_EmpLedgerID As Integer
        Public Str_Type As String
        Public Str_Description As String
        Public dbl_AmountRequested As Double
        Public Str_Currency As String
        Public dbl_ApprovedAmount As Double
        Public Str_Status As String
        Public dbl_MonthlyDeduct As Double
        Public int_DeductMonthCount As Integer
        Public dtp_StartDate As Date
        Public bool_Approved As Boolean
        Public str_CreatedBy As String
        Public dtp_CreatedDate As DateTime
        Public str_LastUpdatedBy As String
        Public dtp_LastUpdatedDate As DateTime
        Public str_ApprovedBy As String
        Public dtp_ApprovedDate As DateTime
    End Class
    Public Class csAdvanceRequestSub
        Public ReqID As String
        Public VouType As String
        Public VouNo As String
        Public TCDebit As Double
        Public TCCredit As Double
        Public LCDebit As Double
        Public LCCredit As Double
    End Class
End Class
