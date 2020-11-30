Public Class csItemCost
    Inherits csSignature
    Public str_SiteID As String
    Public str_MenuID As String
    Public ObjItemCostMain As New csItemCostMain
    Public ObjItemCostSub As New csItemCostSub
    
    Public Function DBTemplate() As DataTable
        Dim DT_Template As New DataTable
        DT_Template.Columns.Add(New DataColumn("SlNo", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("Alias1"))
        DT_Template.Columns.Add(New DataColumn("Alias2"))
        DT_Template.Columns.Add(New DataColumn("ItemCode"))
        DT_Template.Columns.Add(New DataColumn("CostType"))
        DT_Template.Columns.Add(New DataColumn("OldPrice", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("NewPrice", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Stock", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Comment"))

        Return DT_Template
    End Function
End Class

Public Class csItemCostMain
    Public int_BusinessPeriodID As Integer
    Public str_Flag As String
    Public str_ItemCostPrefix As String
    Public str_DocumentNo As String
    Public dtp_DocumentDate As Date
    Public str_Comment As String
    Public int_RevNo As Integer
    Public int_LanguageCode As Integer
    Public str_WHID As String
End Class

Public Class csItemCostSub
    Public dt_CostTypeSub As DataTable
End Class

