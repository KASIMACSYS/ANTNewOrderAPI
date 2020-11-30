
Public Class csEmpAtt
    Inherits csSignature
    Public str_SiteID As String
    Public str_Flag As String
    Public int_BusinessPeriodID As Integer

    Public int_LedgerID As Integer
    Public date_AttDate As Date


    Public dt_EmpAtten As DataTable
    Public dt_EmpAttSub As DataTable
    Public dt_Sub As DataTable
    'Public int_Attendance As Integer
    'Public str_Comment As String
    'Public str_TIN1 As String
    'Public str_TIN2 As String
    'Public str_TOUT1 As String
    'Public str_TOUT2 As String
    'Public dbl_OT As Double

    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        '' If CustomerSetting.Item("Project").ToString = "True" Then
        'objProject = New csProjectDetail
        '' End If
    End Sub



    Public Function DBTemplate() As DataTable
        Dim DT_Template As New DataTable
        DT_Template.Columns.Add(New DataColumn("LedgerID", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("Attendance", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("ProjectID"))
        DT_Template.Columns.Add(New DataColumn("Description"))
        DT_Template.Columns.Add(New DataColumn("Location"))
        DT_Template.Columns.Add(New DataColumn("TIN1", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("TOUT1"))
        DT_Template.Columns.Add(New DataColumn("TIN2"))
        DT_Template.Columns.Add(New DataColumn("TOUT2"))
        DT_Template.Columns.Add(New DataColumn("OT", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Extra1"))
        DT_Template.Columns.Add(New DataColumn("Extra2"))
        DT_Template.Columns.Add(New DataColumn("Comment"))
        DT_Template.Columns.Add(New DataColumn("Extra3"))
        DT_Template.Columns.Add(New DataColumn("Extra4", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("Extra5", System.Type.GetType("System.Double")))
        Return DT_Template
    End Function

End Class


Public Class csEmployeeVacation
    Inherits csSignature

    Public str_SiteID As String
    Public str_Flag As String
    Public str_FormPrefix As String
    Public int_BusinessPeriodID As Integer

    Public int_LedgerID As Integer
    Public date_PostDate As Date
    Public date_FromDate As Date
    Public date_ToDate As Date
    Public date_RtnDate As Date
    Public str_Comment As String
    Public str_VacationType As String
End Class