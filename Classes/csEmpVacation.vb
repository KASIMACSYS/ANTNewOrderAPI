Public Class csEmpVacation
    Inherits csSignature

    Public int_CID As Integer
    Public str_Flag As String
    Public int_BusinessPeriodID As Integer
    Public str_VacDetails As String
    Public dtp_From As Date
    Public dtp_To As Date
    Public dtp_Today As Date
    Public str_Comment As String
    Public dtp_RtnDate As Date
    Public int_LedgerID As Integer
    Public str_RtnComment As String
    Public dt_EmpVacation As New DataTable

    Public objProject As csProjectDetail

    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        'If CustomerSetting.Item("Project").ToString = "True" Then
        objProject = New csProjectDetail
        'End If
    End Sub

End Class


