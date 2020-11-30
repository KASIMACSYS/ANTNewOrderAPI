'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports System.Data

Public Class csFormSetting
    Public str_SiteID As String
    Public str_BusinessPerionID As Integer
    Public dt As New DataTable
    Public MenuID As String
    Public MenuName As String
    Public dt_DefaultForm As DataTable
    Public dt_PropertyMenuID As DataTable
    Public dt_GridMenuID As DataTable
    Public dt_GridSetting As DataTable
    Public dt_ReportMenuID As DataTable
    Public dt_ReportSetting As DataTable
    Public dt_PropertySubMenuID As DataTable
    Public dt_PropertySub As DataTable
    Public dt_ApprovalSetting As DataTable
    Public dt_PropertySortViewMenuID As DataTable
    Public dt_PropertySortView As DataTable
    Public str_Flag As String
    Public objproject As csProjectDetail
    Public Flag As String
    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        'If CustomerSetting.Item("useProject").ToString = "True" Then
        objproject = New csProjectDetail
        'End If
    End Sub
End Class
