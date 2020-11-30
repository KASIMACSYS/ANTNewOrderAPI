Public Class csProduction
    Inherits csSignature
    Public str_SiteID As String
    Public ObjProductionMain As New csProductionMain
    Public ObjProductionSub As New csProductionSub
    Public DTBatch As DataTable

    Public Function DBTemplate() As DataTable
        Dim DT_Template As New DataTable
        DT_Template.Columns.Add(New DataColumn("SlNo", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("Alias1"))
        DT_Template.Columns.Add(New DataColumn("Alias2"))
        DT_Template.Columns.Add(New DataColumn("ItemCode"))
        DT_Template.Columns.Add(New DataColumn("Unit"))
        DT_Template.Columns.Add(New DataColumn("BaseUnit", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("VouQty", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("PrimaryQty", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Price", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("BaseUnitPrice", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("TCDiscountAmount", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("LCAmount", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Comment"))
        DT_Template.Columns.Add(New DataColumn("OrgRefNo", System.Type.GetType("System.Int32")))
       
        'DT_Template.Columns("SlNo").AutoIncrement = True
        'DT_Template.Columns("SlNo").AutoIncrementStep = 1
        'DT_Template.Columns("SlNo").AutoIncrementSeed = 1
        Return DT_Template
    End Function
End Class

Public Class csProductionMain
    Public int_BusinessPeriodID As Integer
    Public str_MenuID As String
    Public str_Flag As String
    Public str_Prefix As String
    Public str_ProdNo As String
    Public str_JOBNo As String
    Public dtp_VouDate As Date
    Public dtp_EstDate As Date
    Public dtp_CompDate As Date
    Public str_Status As String
    Public str_Comment As String
    Public str_WHID As String
    Public str_ProdunitName As String
    Public str_LpoNo As String

    Public dbl_TCAmount As Double
    Public dbl_TCDisAmount As Double
    Public dbl_TCDiscountAmount As Double
    Public dbl_TCNetAmount As Double
    Public dbl_TCVatAmount As Double
End Class

Public Class csProductionSub
    Public dt_FGItems As DataTable
End Class
