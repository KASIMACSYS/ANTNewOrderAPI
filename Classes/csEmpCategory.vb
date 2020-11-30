Public Class csEmpCategory
    Public str_SiteID As String
    Public int_BusinessPeriodID As String
    Public str_flag As String
    Public str_Category As String
    Public str_LeaveTypes As String
    Public int_Days As String
    Public dt_empcategory As DataTable
    Public dt_empcategorysub As DataTable
    Public dt_editempcategorysub As DataTable
    Public dt_Gratuity As DataTable
    Public dt_LeaveType As DataTable

    Public Function DBTemplate() As DataTable
        Dim DB_Template As New DataTable
        DB_Template = New DataTable
        DB_Template.Columns.Add(New DataColumn("Category"))
        DB_Template.Columns.Add(New DataColumn("LeaveType"))
        'DB_Template.Columns.Add(New DataColumn("MonthinDays", GetType(Integer)))
        DB_Template.Columns.Add(New DataColumn("DaysinMonth", GetType(Integer)))
        DB_Template.Columns.Add(New DataColumn("LeaveSalary"))
        DB_Template.Columns.Add(New DataColumn("Gratuity"))
        DB_Template.Columns.Add(New DataColumn("DaysinYear", GetType(Integer)))
        DB_Template.Columns.Add(New DataColumn("VacationDays", GetType(Integer)))
        DB_Template.Columns.Add(New DataColumn("CarryForward", GetType(Integer)))
        DB_Template.Columns.Add(New DataColumn("PassageAmount", GetType(Decimal)))
        Return DB_Template
    End Function
    Public Function DBCategorysub() As DataTable
        Dim DB_CategorySub As New DataTable
        DB_CategorySub = New DataTable

        DB_CategorySub.Columns.Add(New DataColumn("Category"))
        DB_CategorySub.Columns.Add(New DataColumn("LeaveDetails"))
        DB_CategorySub.Columns.Add(New DataColumn("IsApplicable", GetType(Boolean)))
        Return DB_CategorySub
    End Function

    Public Function DBGratuity() As DataTable
        Dim DB_Gratuity As New DataTable
        DB_Gratuity = New DataTable
        DB_Gratuity.Columns.Add(New DataColumn("Type"))
        DB_Gratuity.Columns.Add(New DataColumn("Category"))
        DB_Gratuity.Columns.Add(New DataColumn("FromVal", GetType(Decimal)))
        DB_Gratuity.Columns.Add(New DataColumn("ToVal", GetType(Decimal)))
        DB_Gratuity.Columns.Add(New DataColumn("Calculation", GetType(Decimal)))
        Return DB_Gratuity
    End Function

End Class
