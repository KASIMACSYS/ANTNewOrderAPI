'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class csUserMgt

    Public int_SiteID As Integer
    Public int_UserID_Old As Integer
    Public int_UserID As Integer
    Public str_UserName As String
    Public str_Password As String
    Public str_GroupID As String
    Public bool_InActive As Boolean
    Public str_DefaultSiteID As String
    Public int_EmployeeLedgerID As Integer
    Public str_HeaderandButtonBackColor As String
    Public str_FormbackColor As String

    Public int_CreatedBy As Integer
    Public dtp_CreatedDate As Date
    Public int_LastUpdatedBy As Integer
    Public dtp_LastUpdatedDate As Date
    Public bool_ShowPopUp As Boolean

    Public dt_UserMgt As DataTable
    Public dt_UserDynSettings As DataTable
    Public dt_UserMain As DataTable
    Public str_Flag As String

    Public str_SalesManID As String

    Public int_LanguageCode As Integer
    Public str_ActiveDirectoryPath As String
    Public str_ActiveDirectoryDomain As String
    Public str_ActiveDirectoryUserID As String
End Class
Public Class csUserPriceList

    Public int_CID As Integer
    Public Int_SalesMan As Integer
    Public Int_UserID As Integer
    Public str_Flag As String
    Public Str_CreatedBy As String
    Public dtp_CreatedDate As Date
    Public Str_LastUpdatedBy As String
    Public dtp_LastUpdatedDate As Date
    Public Str_ApprovedBy As String
    Public dtp_ApprovedDate As Date
    Public dt_UserPriceList As DataTable
End Class
