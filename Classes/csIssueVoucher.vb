'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class csIssueVoucher
    Inherits csSignature
    Public int_CID As Integer
    Public ObjIssueVoucherMain As New csIssueVoucherMain
    Public ObjIssueVoucherSub As New csIssueVoucherSub
    Public objproject As csProjectDetail
    Public DTBatch As New DataTable

    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        '' If CustomerSetting.Item("Project").ToString = "True" Then
        objproject = New csProjectDetail
        '' End If
    End Sub

    Public Function DBTemplate() As DataTable
        Dim DT_Template As New DataTable
        DT_Template = New DataTable
        DT_Template.Columns.Add(New DataColumn("Slno"))
        DT_Template.Columns.Add(New DataColumn("Alias1"))
        DT_Template.Columns.Add(New DataColumn("Alias2"))
        DT_Template.Columns.Add(New DataColumn("ItemCode"))
        DT_Template.Columns.Add(New DataColumn("Unit"))
        DT_Template.Columns.Add(New DataColumn("BaseUnit", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Desc1"))
        DT_Template.Columns.Add(New DataColumn("Desc2"))
        DT_Template.Columns.Add(New DataColumn("Desc3"))
        DT_Template.Columns.Add(New DataColumn("LpoNo"))
        DT_Template.Columns.Add(New DataColumn("VouQty", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("PrimaryQty", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Wac", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("BaseUnitPrice", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Returned", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Damaged", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Used", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("PriceType"))

        'DT_Template.Columns("Wac").DefaultValue = 0
        'DT_Template.Columns("VouQty").DefaultValue = 0
        'DT_Template.Columns("PrimaryQty").DefaultValue = 0
        'DT_Template.Columns("Returned").DefaultValue = 0
        'DT_Template.Columns("Damaged").DefaultValue = 0
        'DT_Template.Columns("Used").DefaultValue = 0

        Return DT_Template
    End Function
End Class

Public Class csIssueVoucherMain
    Public str_VouNo As String
    Public str_MenuID As String
    Public str_JONo As String
    Public int_LedgerID As Integer
    Public str_Alias As String
    Public dtp_VouDate As Date
    Public dtp_ReturnDate As Date
    Public str_Comment As String
    Public bool_Status As Boolean
    Public str_ReqFormNo As String
    Public int_RevNo As Integer
    Public str_IssuedBy As String
    Public str_ReturnedBy As String
    Public str_Payterm As String
    Public int_BusinessPeriodID As Integer
    Public bool_ApprovedStatus As Boolean
    Public str_Flag As String
    Public str_IssueVoucherPrefix As String
    Public str_ProductionUnitNo As String
    Public int_LanguageCode As Integer
    Public str_WHID As String
    Public str_DstLedger As String
    Public str_DstLedgerDesc As String
End Class

Public Class csIssueVoucherSub
    Public dt_IssueVoucherSub As DataTable
End Class



