'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class csCurrencyMaster
    Public str_SiteID As String
    Public str_CurrencyID As String
    Public str_CurrencyCode As String
    Public str_CurrencyName As String
    Public dbl_PurExchangeRate As Double
    Public dbl_ExchangeRate As Double
    Public dbl_SalExchangeRate As Double
    Public bool_DefaultCurrency As Boolean
    Public int_DecimalPlace As Integer
    Public str_MajorCurrency As String
    Public str_MinorCurrency As String

    Public str_CreatedBy As String
    Public dtp_CreatedDate As DateTime
    Public str_LastUpdatedBy As String
    Public dtp_LastUpdatedDate As DateTime
    Public int_BusinessPeriodID As Integer
    Public str_ApprovedBy As String
    Public dtp_ApprovedDate As Date
    Public bool_ApprovedStatus As Boolean
    Public str_Flag As String

End Class
