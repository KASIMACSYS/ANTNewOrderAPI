
Public Class csEmpPayslip
    Inherits csSignature
    Public str_SiteID As String
    Public str_CurrencyCode As String
    Public objPSMain As New csEmp_PayslipMain

    Public Function PaySlipMain() As DataTable
        Dim DT_PaySlipMain As New DataTable

        DT_PaySlipMain.Columns.Add(New DataColumn("PSNo"))
        DT_PaySlipMain.Columns.Add(New DataColumn("LedgerID", System.Type.GetType("System.Int32")))
        DT_PaySlipMain.Columns.Add(New DataColumn("WorkDays", System.Type.GetType("System.Double")))
        DT_PaySlipMain.Columns.Add(New DataColumn("PresDays", System.Type.GetType("System.Double")))
        DT_PaySlipMain.Columns.Add(New DataColumn("AbsDays", System.Type.GetType("System.Double")))
        'DT_PaySlipMain.Columns.Add(New DataColumn("NotDeductAbsDaysAmt", System.Type.GetType("System.Double")))
        DT_PaySlipMain.Columns.Add(New DataColumn("OTHours", System.Type.GetType("System.Double")))
        'DT_PaySlipMain.Columns.Add(New DataColumn("HoliOTHours", System.Type.GetType("System.Double")))
        DT_PaySlipMain.Columns.Add(New DataColumn("Earnings", System.Type.GetType("System.Decimal")))
        DT_PaySlipMain.Columns.Add(New DataColumn("Deductions", System.Type.GetType("System.Decimal")))
        DT_PaySlipMain.Columns.Add(New DataColumn("NetAmount", System.Type.GetType("System.Decimal")))
        DT_PaySlipMain.Columns.Add(New DataColumn("Comment"))
        DT_PaySlipMain.Columns.Add(New DataColumn("PaymentStatus"))
        'DT_PaySlipMain.Columns.Add(New DataColumn("ESHSlNo"))
        'DT_PaySlipMain.Columns.Add(New DataColumn("IsChanged"))
        Return DT_PaySlipMain
    End Function

    Public Function PaySlipSub() As DataTable
        Dim DT_PaySlipSub As New DataTable
        DT_PaySlipSub.Columns.Add(New DataColumn("PSNo"))
        DT_PaySlipSub.Columns.Add(New DataColumn("LedgerID"))
        DT_PaySlipSub.Columns.Add(New DataColumn("Tag"))
        DT_PaySlipSub.Columns.Add(New DataColumn("Value", System.Type.GetType("System.Decimal")))
        Return DT_PaySlipSub
    End Function

    Public Function PaySlipParam() As DataTable
        Dim DT_PaySlipParam As New DataTable
        DT_PaySlipParam.Columns.Add(New DataColumn("Tag"))
        'DT_PaySlipParam.Columns.Add(New DataColumn("Parameter"))
        DT_PaySlipParam.Columns.Add(New DataColumn("Value", System.Type.GetType("System.Decimal")))
        Return DT_PaySlipParam
    End Function
End Class

Public Class csEmp_PayslipMain
    Public str_Flag As String
    Public str_FormPrefix As String
    Public str_MenuID As String
    Public str_Type As String
    Public int_BusinessPeriodID As Integer

    Public str_PSRef As String
    Public str_PSNo As String
    Public int_RevNo As Integer
    Public date_PSMonth As Date
    Public date_PostDate As Date

    Public DT_PaySlipMain As DataTable
    Public DT_PaySlipSub As DataTable
    Public DT_PayslipParam As DataTable
    Public dt_InvoiceAccounts As DataTable
    Public dt_Wages As New DataTable
    Public dt_InvoiceAccountsCostCentre As New DataTable
End Class


Public Class csEmpPaySlipMain
    Public str_SiteID As String
    Public str_BusinessPerionID As Integer
    Public dt_Main As New DataTable
    Public dtp_FromDate As Date
    Public dtp_ToDate As Date
    Public dtp_Date As String
    Public str_EmpName As String
    Public str_Categotry As String
    Public str_VouNo As String
End Class