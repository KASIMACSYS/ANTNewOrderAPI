'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports System.Data
Public Class csMainForms
    Public str_SiteID As String
    Public str_BusinessPerionID As Integer
    Public dt_Main As New DataTable
    Public MenuID As String
    Public str_MerchantName As String = String.Empty
    Public str_SalesManID As String
    Public bool_Status As Boolean
    Public dtp_FromDate As Date
    Public dtp_ToDate As Date
    Public dtp_Date As String
    Public str_CurrencyCode As String
    Public str_Project As String
    Public bool_Open As Boolean
    Public bool_Closed As Boolean
    Public bool_Partial As Boolean
    Public bool_Manually As Boolean
    Public bool_Invoiced As Boolean
    Public bool_Uninvoiced As Boolean
    Public bool_Cancelled As Boolean
    Public bool_Draft As Boolean
    Public bool_Paid As Boolean = False
    Public bool_NotPaid As Boolean = False
    Public bool_PartiallyPaid As Boolean = False
    Public str_ApprovedStatus As String
    Public int_AccountingPeriodFrom As Integer = 0
    Public str_AccountingPeriod As String
    Public int_AccountingPeriodTo As Integer = 0
    Public str_WorkOrderNo As String = String.Empty
    Public str_CreatedBy As String
    Public str_WareHouse As String
    Public str_SignatureType As String
    Public str_WHMaster As String
    Public str_DateType As String
    Public str_User As String
    Public int_GrpID As Integer
    Public str_LedgerDept As String
    Public str_Options As String
    Public str_GrpName As String
    Public bool_Options2 As Boolean
    Public bool_Options3 As String = String.Empty
    Public str_Filter1 As String
    Public str_Filter2 As String
    Public bool_QtyStatusAll As Boolean
    Public bool_QtyStatusOpen As Boolean
    Public bool_QtyStatusPartial As Boolean
    Public bool_QtyStatusClose As Boolean
End Class
