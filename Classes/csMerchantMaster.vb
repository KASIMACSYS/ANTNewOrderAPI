'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Public Class csMerchantMaster
    Public str_SiteID As String
    Public int_BusinessPeriodID As Integer
    Public ObjMerSub As New csMerchantSub
    Public ObjMerMain As New csMerMain
    'Public ObjMerLedger As New csMerLedger
End Class

Public Class csMerMain
    Public str_LedgerID As String
    Public int_SalesAccLedgerID As Integer
    Public int_PurchaseAccLedgerID As Integer
    Public int_CashSalesAccLedgerID As Integer
    Public int_CashPurchaseAccLedgerID As Integer
    Public int_SalesRTNAccLedgerID As Integer
    Public int_PurchaseRTNAccLedgerID As Integer
    Public str_LedgerDesc As String
    Public str_MerchantID As String
    Public str_MerchantName As String
    Public str_Type As String
    Public str_Alias1 As String
    Public str_Alias2 As String
    Public str_Trn As String
    Public str_Consignee As String
    Public bool_SendSMS As Boolean
    Public bool_SendEmail As Boolean

    Public bool_CusActiveStatus As Boolean
    Public int_CusCreditDays As Integer
    Public str_CusPayTerm As String
    Public dbl_CusCreditLimitAmount As Double
    Public int_CusCreditLimitCondition As Integer
    Public int_CusAgingLimitCondition As Integer
    Public bool_CusCreditDaysRemindExpiry As Boolean
    Public bool_CusCreditDaysPDC As Boolean
    Public bool_CusCreditLimitPDC As Boolean

    Public bool_VenActiveStatus As Boolean
    Public int_VenCreditDays As Integer
    Public str_VenPayTerm As String
    Public dbl_VenCreditLimitAmount As Double
    Public int_VenCreditLimitCondition As Integer
    Public int_VenAgingLimitCondition As Integer
    Public bool_VenCreditDaysRemindExpiry As Boolean
    Public bool_VenCreditDaysPDC As Boolean

    Public bool_ReverseCharge As Boolean

    Public bool_IsSellingPercentage As Boolean
    Public int_SellingPercentage As Integer
    Public str_Contact As String
    Public str_Address As String
    Public str_DelivAddress As String
    Public dtp_MerchantSince As Date
    Public str_PoBox As String
    Public str_Tel As String
    Public str_Mobile As String
    Public str_Fax As String
    Public str_Email As String
    Public str_Comment As String
    Public str_PopUpComment As String
    Public str_ChequePrintingName As String
    Public bool_IntraCompanyFlag As Boolean

    Public str_CreatedBy As String
    Public dtp_CreatedDate As DateTime
    Public str_LastUpdatedBy As String
    Public dtp_LastUpdatedDate As DateTime
    Public str_DefaultItemValue As String
    Public int_SalesMan As Integer
    Public str_Flag As String
    Public dtp_date As Date
    Public str_City As String
    Public str_Filter1 As String
    Public str_Filter2 As String
    Public str_Filter3 As String
    Public str_Filter4 As String
    Public str_ItemDiscType As String
    Public str_Desc1 As String
    Public str_Desc2 As String
    Public str_Desc3 As String
    Public str_Desc4 As String
    Public str_Desc5 As String
    Public str_Desc6 As String
    Public str_Desc7 As String
    Public str_Desc8 As String
    Public str_TimeZone As String
    Public str_SupportMailID As String
    Public str_Country As String
    Public str_Region As String
End Class

Public Class csMerchantSub
    Public STSiteID As String
End Class




