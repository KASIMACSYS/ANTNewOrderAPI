'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports System.Data
Public Class csCustomizedItemReport
    Public str_SiteID As String
    Public str_BusinessPerionID As Integer
    Public dt_Main As DataTable
    Public ItemCode As String
    Public dtp_FromDate As Date
    Public dtp_ToDate As Date
    Public str_MerchantID As String
    Public FromRange As String
    Public ToRange As String
    Public dtp_Date As String
    Public dt_IssueVou As DataTable
    Public WHLocation As String
    Public str_Flag As String
    Public dt_ItemCode As DataTable
    Public objCustomizedVoucher As New csCustomizedVoucherSearch
    Public bool_ZeroSuppress As Boolean

End Class
Public Class csCustomizedVoucherSearch
    Public bool_QTN As Boolean
    Public bool_SalOrd As Boolean
    Public bool_DO As Boolean
    Public bool_SIS As Boolean
    Public bool_ISSUE As Boolean
    Public bool_GIP As Boolean
    Public bool_SRT As Boolean
    Public bool_LPO As Boolean
    Public bool_MRV As Boolean
    Public bool_PIP As Boolean
    Public bool_PRT As Boolean
    Public bool_GEP As Boolean
    Public bool_RVCash As Boolean
    Public bool_RVCheq As Boolean
    Public bool_PVCash As Boolean
    Public bool_PVCheq As Boolean
    Public bool_GEV As Boolean
    Public bool_JV As Boolean
    Public bool_PRODUCTION As Boolean
    Public str_SalesMan As String
    Public str_WildSerach As String
End Class
