'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Public Class csSIS
    Public str_SiteID As String
    Public objSISMain As New csSISMain
    Public objSISSub As New csSISSub
    Public objSISProject As New csSISProject
    Public objSISCommon As New csSISCommon
End Class

Public Class csSISMain
    Public str_SISNo As String
    Public str_MerchantID As String
    Public str_MerchantName As String
    Public str_InvRef As String
    Public dtp_InvDate As Date
    Public dtp_DueDate As Date
    Public int_Aging As Integer
    Public str_PayTerm As String

    Public dbl_TCTotalAmount As Double
    Public dbl_TCDiscountAmount As Double
    Public dbl_TCMiscAmount As Double
    Public dbl_TCTaxAmount As Double
    Public dbl_TCDisAmount As Double
    Public dbl_TCNetAmount As Double
    Public dbl_TCPDCAmount As Double

    Public bool_CounterSales As Boolean
    Public bool_AffectInventory As Boolean
    Public str_PaymentStatus As String
    Public int_RevNo As Integer
    Public str_Comment As String

    Public dbl_LCNetAmount As Double
    Public dbl_LCPDCAmount As Double

    Public dbl_ExchangeRate As Double
    Public str_CurrencyID As String

    Public str_CreatedBy As String
    Public dtp_CreatedDate As Date
    Public str_LastUpdatedBy As String
    Public dtp_LastUpdatedDate As Date
    Public int_BusinessPeriodID As Integer
    Public str_ApprovedBy As String
    Public dtp_ApprovedDate As Date
    Public bool_ApprovedStatus As Boolean
    Public bool_DirectInvoice As Boolean
End Class

Public Class csSISSub
    Public dt_SISMatching As DataTable
End Class

Public Class csSISProject
    Public str_ProjectID As String
    Public str_WorkOrderNo As String
    Public str_ProjectLocation As String
End Class

Public Class csSISCommon
    Public dt_Generic As DataTable
    Public str_Flag As String
    Public str_SISPrefix As String
End Class



