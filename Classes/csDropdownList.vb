Public Class csDropdownList
    Inherits csSignature
    Public str_SiteID As String
    Public int_BusinessPerionID As Integer
    Public str_Types As String
    Public str_Flag As String
    Public dt_BaseType As DataTable
    Public dt_salarySlap As DataTable
    Public dt_EditSlap As DataTable
    Public dt_TempBaseType As DataTable
    Public Function DBSalaryParticulars() As DataTable
        Dim DB_Salaryparticular As New DataTable
        DB_Salaryparticular = New DataTable
        DB_Salaryparticular.Columns.Add(New DataColumn("Particulars"))
        DB_Salaryparticular.Columns.Add(New DataColumn("CustomText"))
        DB_Salaryparticular.Columns.Add(New DataColumn("CalcType"))
        DB_Salaryparticular.Columns.Add(New DataColumn("IsAffectNetSalary", GetType(Boolean)))
        DB_Salaryparticular.Columns.Add(New DataColumn("IsInLeaveSalary", GetType(Boolean)))
        DB_Salaryparticular.Columns.Add(New DataColumn("IsInGratuity", GetType(Boolean)))
        DB_Salaryparticular.Columns.Add(New DataColumn("IsDeductable", GetType(Boolean)))
        DB_Salaryparticular.Columns.Add(New DataColumn("IsVisible", GetType(Boolean)))
        Return DB_Salaryparticular
    End Function

    Public Function DBSalarySlap() As DataTable
        Dim DB_Salaryslap As New DataTable
        DB_Salaryslap = New DataTable
        DB_Salaryslap.Columns.Add(New DataColumn("Tag"))
        DB_Salaryslap.Columns.Add(New DataColumn("FromSlap", GetType(Decimal)))
        DB_Salaryslap.Columns.Add(New DataColumn("ToSlap", GetType(Decimal)))
        DB_Salaryslap.Columns.Add(New DataColumn("SlapType"))
        DB_Salaryslap.Columns.Add(New DataColumn("ValueBasis", GetType(Decimal)))
        Return DB_Salaryslap
    End Function
End Class
