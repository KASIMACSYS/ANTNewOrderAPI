'======================================================================================
'$Author: Kasim $
'$Rev: 510 $
'$Date: 2013-12-11 11:35:21 +0530 (Wed, 11 Dec 2013) $ 
'======================================================================================
Public Class csBOM
    Inherits csSignature

    Public str_CID As String

    Public objBOMMain As New BOMMain
    Public DT_BOMItemDetails As New DataTable
    Public DT_BOMParameters As New DataTable

    Public Function DT_BOMItemDetailsTemplate() As DataTable
        DT_BOMItemDetailsTemplate = New DataTable
        DT_BOMItemDetailsTemplate.Columns.Add(New DataColumn("SlNo", System.Type.GetType("System.Int32")))
        DT_BOMItemDetailsTemplate.Columns.Add(New DataColumn("ItemCode"))
        DT_BOMItemDetailsTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_BOMItemDetailsTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_BOMItemDetailsTemplate.Columns.Add(New DataColumn("Unit"))
        DT_BOMItemDetailsTemplate.Columns.Add(New DataColumn("UseFormula", GetType(Boolean)))
        DT_BOMItemDetailsTemplate.Columns.Add(New DataColumn("Formula"))
        DT_BOMItemDetailsTemplate.Columns.Add(New DataColumn("Qty", System.Type.GetType("System.Double")))
        Return DT_BOMItemDetailsTemplate
    End Function
End Class

Public Class BOMMain
    Public str_Flag As String
    Public str_FormPrefix As String
    Public str_MenuID As String
    Public int_BusinessPeriodID As Integer
    Public dtp_Date As Date
    Public int_RevNo As Integer
    Public str_BOMNo As String
    Public str_BOMDesc As String
    Public str_ItemCode As String
    Public str_ItemDesc As String
    Public str_Comment As String
End Class