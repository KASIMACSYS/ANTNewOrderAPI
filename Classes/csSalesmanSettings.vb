Public Class csEmployeeSettings
    Inherits csSignature

    Public int_CID As Integer
    Public int_BusinessPeriodID As Integer
    Public str_Flag As String

    Public str_SalesManID As String
    Public str_SalesManName As String
    Public str_Alias1 As String
    Public str_Alias2 As String
    Public str_EmployeeLedgerID As String
    Public dbl_PaymentLimit As Double
    Public int_LimitStatus As Integer
    Public dbl_Commission As Double
    Public str_Comment As String
    Public bool_InActive As Boolean

    Public dt_EmployeeSettings As New DataTable
    Public objProject As csProjectDetail

    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        'If CustomerSetting.Item("Project").ToString = "True" Then
        objProject = New csProjectDetail
        'End If
    End Sub

End Class
