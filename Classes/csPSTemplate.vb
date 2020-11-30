Public Class csPSTemplate
    Inherits csSignature

    Public int_CID As Integer
    Public str_Flag As String
    Public str_VouNo As String
    Public int_RevNo As Integer
    Public str_MenuID As String
    Public str_Prefix As String
    Public str_Interval As String
    Public str_Description As String
    Public str_Comment As String
    Public dt_PSTemplatesub As DataTable

    Public Function DBTemplate() As DataTable
        Dim DT_Template As New DataTable
        DT_Template.Columns.Add(New DataColumn("SlNo", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("Tag"))
        DT_Template.Columns.Add(New DataColumn("Formula"))
        DT_Template.Columns.Add(New DataColumn("Parameter", GetType(Boolean)))
        DT_Template.Columns.Add(New DataColumn("AffectNetAmount", GetType(Boolean)))
        DT_Template.Columns.Add(New DataColumn("Visibility", GetType(Boolean)))
        DT_Template.Columns.Add(New DataColumn("Readonly", GetType(Boolean)))
        DT_Template.Columns.Add(New DataColumn("PSMainTag"))
        Return DT_Template
    End Function
End Class
