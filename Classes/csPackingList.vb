'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Public Class csPackingList

    Inherits csSignature
    Public str_SiteID As String
    Public ObjPackListMain As New csPackingListMain
    Public ObjPackListSub As New csPackingListSub
    Public ObjPackListCommon As New csPackListCommon
    Public objMerchantDetails As New csCustomerDetails
    Public objProject As csProjectDetail
    Public DTBatch As New DataTable
    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        'If CustomerSetting.Item("Project").ToString = "True" Then
        objProject = New csProjectDetail
        'End If
    End Sub


    Public Function DT_PackingListTemplate() As DataTable
        DT_PackingListTemplate = New DataTable
        DT_PackingListTemplate.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
        DT_PackingListTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_PackingListTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_PackingListTemplate.Columns.Add(New DataColumn("ItemCode"))
        DT_PackingListTemplate.Columns.Add(New DataColumn("Unit"))
        DT_PackingListTemplate.Columns.Add(New DataColumn("VouQty", System.Type.GetType("System.Double")))
        DT_PackingListTemplate.Columns.Add(New DataColumn("Comment"))
        DT_PackingListTemplate.Columns.Add(New DataColumn("Desc1"))
        DT_PackingListTemplate.Columns.Add(New DataColumn("Desc2"))
        DT_PackingListTemplate.Columns.Add(New DataColumn("Desc3"))
        DT_PackingListTemplate.Columns.Add(New DataColumn("Desc4"))
        DT_PackingListTemplate.Columns.Add(New DataColumn("Desc5"))
        DT_PackingListTemplate.Columns.Add(New DataColumn("Desc6"))
        DT_PackingListTemplate.Columns.Add(New DataColumn("Desc7"))
        DT_PackingListTemplate.Columns.Add(New DataColumn("Desc8"))
        Return DT_PackingListTemplate
    End Function
End Class

Public Class csPackingListMain
    Public str_DoNo As String
    Public str_InvNo As String

    Public str_PkNo As String
    Public dtp_EntryDate As Date
    Public dtp_DocumentDate As Date
    Public str_Comment As String
    Public str_CreatedBy As String
    Public dtp_CreatedDate As Date
    Public str_LastUpdatedBy As String
    Public dtp_LastUpdatedDate As Date
    Public int_BusinessPeriodID As Integer
    Public dtp_DODate1 As Date
    Public dtp_DoDate2 As Date
    Public str_WHID As String
    Public int_RevNo As Integer
    Public int_Aging As Integer
    Public str_Alias As String
    Public str_PayTerm As String
    Public int_StatusCancel As Integer
    Public str_TCCurrency As String
    Public dbl_ExchangeRate As Double
    Public str_SalesManID As String
    Public str_SalesManName As String
    Public str_PackagingComment As String

    Public str_Desc1 As String
    Public str_Desc2 As String
    Public str_Desc3 As String
    Public str_Desc4 As String
    Public str_Desc5 As String
    Public str_Desc6 As String
    Public str_Desc7 As String
    Public str_Desc8 As String
End Class

Public Class csPackingListSub
    Public dt_PackListSub As DataTable
    Public dt_Sub As DataTable
    Public MenuID As String
    Public str_PkNo As String
    Public str_Flag As String
End Class

Public Class csPackListCommon
    Public str_Flag As String
    Public str_FormPrefix As String
    Public int_LedgerID As Integer
End Class


