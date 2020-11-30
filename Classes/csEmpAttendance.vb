Public Class csEmpAttendance
    Inherits csSignature
    Public str_SiteID As String
    Public str_Flag As String
    Public int_BusinessPeriodID As Integer
    Public int_LedgerID As Integer
    Public date_AttDate As Date
    Public dt_EmpAtten As DataTable
    Public dt_EmpAttSub As DataTable
    Public dt_Sub As DataTable

    Public Function DBTemplate() As DataTable
        Dim DT_Template As New DataTable
        DT_Template.Columns.Add(New DataColumn("LedgerID", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("ID"))
        DT_Template.Columns.Add(New DataColumn("Name"))
        DT_Template.Columns.Add(New DataColumn("Attendance", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("OT", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("TotalHours", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Comment"))
        DT_Template.Columns.Add(New DataColumn("Category"))
        Return DT_Template
    End Function

End Class

