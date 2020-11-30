Public Class csTaxFileReturn
    Inherits csSignature

    Public str_CID As String
    Public str_MenuID As String
    Public str_Flag As String
    Public str_Condition As String
    Public str_VouNo As String
    Public Str_FormPrefix As String
    Public str_Description As String
    Public str_TaxAgentUID As String
    Public dtp_VouDate As Date
    Public dtp_FromDate As Date
    Public dtp_ToDate As Date
    Public str_Status As String
    Public dt_TaxFileReturn As DataTable
    Public bool_All As Boolean = False
    Public bool_Open As Boolean = False
    Public bool_Submitted As Boolean = False
    Public str_VouType As String = String.Empty
    Public int_LedgerID As Integer

End Class
