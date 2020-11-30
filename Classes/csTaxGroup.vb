Public Class csTaxGroup
    Inherits csSignature
    Public str_SiteID As String
    Public int_BusinessPeriodID As Integer
    Public ObjTaxGroupMain As New csTaxGroupMain
    Public ObjTaxGroupSub As New csTaxGroupSub

    Public Function DTTemplate() As DataTable
        Dim DT_Template As New DataTable
        DT_Template.Columns.Add(New DataColumn("TaxCode"))
        Return DT_Template
    End Function

End Class
Public Class csTaxGroupMain
    Public str_Flag As String
    Public int_TaxGroupID As Integer
    Public str_TaxGroupName As String
    Public str_TaxCode As String
    Public str_Description As String
    Public str_CreatedBy As String
    Public dtp_CreatedDate As Date = Date.Now
    Public str_LastUpdatedBy As String
    Public dtp_LastUpdatedDate As Date = Date.Now
    Public dt_TaxGroup As DataTable
End Class
Public Class csTaxGroupSub
    Public str_TaxCode As String
    Public dt_TaxGroupSub As DataTable
End Class

