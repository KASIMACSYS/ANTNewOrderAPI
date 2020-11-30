'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class csProductGrouping

    Public str_SiteID As String
    Public int_BusinessPerionID As Integer

    Public str_ProductID As String
    Public str_Description As String
    Public str_ParentID As String
    Public str_CreatedBy As String
    Public dtp_CreatedDate As DateTime
    Public str_LastUpdatedBy As String
    Public dtp_LastUpdatedDate As DateTime
    Public dt_ItemCode As DataTable
    Public str_Flag As String
    Public str_groupname As String
    Public str_ItemColumn As String
    Public dt_selecteditem, dt_selectedall As New DataTable
    Public bool_Price As Boolean
    Public Condition As String
End Class

