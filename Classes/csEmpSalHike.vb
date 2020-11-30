
Public Class csEmpSalHike
    Inherits csSignature

    Public str_FormPrefix As String
    Public str_VouNo As String
    Public int_CID As Integer
    Public int_BusinessPeriodID As Integer
    Public str_LedgerID As Integer
    Public str_EmpName As String
    Public dtp_VoucherDate As Date
    Public str_comment As String
    Public str_salaryHike As String
    Public str_AllowHike As String
    Public str_IncenHike As String
    Public str_HraHike As String
    Public str_UtilHike As String
    Public str_SalHike As String
    Public str_Flag As String
    Public dt_EmpDocumnet As DataTable
    Public dt_SalHikeDetails As DataTable
    Public dt_Main As DataTable
    Public dt_SalaryHike As New DataTable
    Public str_Category As String
    Public dt_SalaryArrears As DataTable
    Public str_Tag As String = String.Empty
    Public dbl_Percentage As Decimal
End Class


