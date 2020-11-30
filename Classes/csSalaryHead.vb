Public Class csSalaryHead
    Inherits csSignature
    Public str_SiteID As String
    Public int_BusinessPerionID As Integer
    Public str_Flag As String
    Public dt_Salaryhead As DataTable
    Public dt_Source As DataTable
    Public Function DBSalaryHead() As DataTable
        Dim DB_SalaryHead As New DataTable
        DB_SalaryHead = New DataTable
        DB_SalaryHead.Columns.Add(New DataColumn("Code"))
        DB_SalaryHead.Columns.Add(New DataColumn("Description"))
        DB_SalaryHead.Columns.Add(New DataColumn("Source"))
        Return DB_SalaryHead
    End Function
End Class
