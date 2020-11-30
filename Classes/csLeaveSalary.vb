
Public Class csLeaveSalary
    Inherits csSignature

    Public _SiteID As String
    Public _Flag As String
    Public _FormPrefix As String
    Public _MenuID As String
    Public _BusinessPeriodID As Integer
    Public _RevNo As Integer

    Public _RefNo As String
    Public _LedgerID As Integer
    Public _PostDate As Date
    Public _FromDate As Date
    Public _ToDate As Date
    Public _NoofDays As Integer
    Public _LeaveDays As Integer
    Public _PerDay As Double
    Public _TotalLeaveDays As Double

    Public _PSAmt As Double
    Public _MonthlyLeaveAmount As Double
    Public _PerDayLeaveAmount As Double
    Public _CalcLeaveAmt As Double
    Public _ChagLeaveAmt As Double

    Public _PassageAmt As Double
    Public _PerDayPassageAmt As Double
    Public _CalcPassageAmt As Double
    Public _ChagPassageAmt As Double

    Public _TotalAmt As Double
    Public _LessAmt As Double
    Public _NetAmt As Double

    Public _Comment As String
    Public _ExcludeEmergency As Boolean

    Public _BasicSalaryforGratuity As Double
    Public _CalcLastMonthSalary As Double
    Public _ChangeLastMonthSalary As Double

    Public _BasicSalary As Double ' for gratuity lable value
    Public _PerDayGratuity As Double
    Public _TotalDaysGratuity As Double
    Public _PensionAmt As Double

    Public _CalcGratuity As Double
    Public _ChangeGratuity As Double
    Public _JoiningDate As Date
    Public _PSMonth As Date
End Class
