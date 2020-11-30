Public Class csBoardProduction

    Inherits csSignature
    Public str_SiteID As String
    Public int_BusinessPeriodID As Integer
    Public ObjProdConsumption As New csProductionConsumption

    Public Class csProductionConsumption
        Public Str_Flag As String
        Public Str_MenuID As String
        Public Str_FormPrefix As String
        Public dtp_VouDate As Date
        Public int_RevNo As Integer
        Public Str_VoucherID As String
        Public Str_Comment As String

        Public XMLData1 As String
        Public XMLData2 As String
        Public XMLData3 As String
    End Class

End Class

Public Class csBoardProduction_Ledger
    Public Str_Flag As String
    Public Str_MenuID As String
    Public str_Datestring As String
    Public dtp_FromDate As Date
    Public dtp_ToDate As Date
    Public bool_All As Boolean
    Public bool_Board As Boolean
    Public bool_Plaster As Boolean
    Public bool_VAP As Boolean
    Public _DTLedger As DataTable
End Class