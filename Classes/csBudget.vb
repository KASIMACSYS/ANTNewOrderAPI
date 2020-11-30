
Public Class csBudget
    Inherits csSignature
    Public str_SiteID As String
    Public int_BusinessPeriodID As Integer
    Public ObjBudgetMain As New csBudgetMain
    Public ObjBudgetSub As New csBudgetSub

    Public Class csBudgetMain
        Public Str_Flag As String
        Public str_GroupLedger As Integer
        Public Str_MenuID As String
        Public Str_FormPrefix As String
        Public dtp_VouDate As Date
        Public int_RevNo As Integer
        Public dtp_FromDate As Date
        Public dtp_ToDate As Date
        Public Str_BudgetID As String
        Public str_Description As String
        Public Str_Comment As String
        Public Str_Status As String
        Public Str_User As String
        Public dt_Budget As DataTable
        Public int_AccountingPeriodFrom As Integer = 0
        Public str_AccountingPeriod As String
        Public dtp_Date As String
        Public str_DateType As String
        Public str_SignatureType As String
    End Class

    Public Class csBudgetSub
        Public BudgetID As String
        Public VouType As String
        Public VouNo As String
        Public TCDebit As Double
        Public TCCredit As Double
        Public LCDebit As Double
        Public LCCredit As Double
    End Class

End Class
