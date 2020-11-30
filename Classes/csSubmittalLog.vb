
Public Class csSubmittalLog
    Inherits csSignature
    Public str_SiteID As String
    Public int_BusinessPeriodID As Integer
    Public ObjSubmittalLogMain As New csSubmittalLogMain
    Public ObjSubmittalLogSub As New csSubmittalLogSub
    Public objproject As New csProjectDetail


    Public Function DT_SubmittalLogTemplate() As DataTable
        DT_SubmittalLogTemplate = New DataTable
        DT_SubmittalLogTemplate.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
        DT_SubmittalLogTemplate.Columns.Add(New DataColumn("DocNo"))
        DT_SubmittalLogTemplate.Columns.Add(New DataColumn("Title"))
        DT_SubmittalLogTemplate.Columns.Add(New DataColumn("Status"))

        DT_SubmittalLogTemplate.Columns.Add(New DataColumn("Ref"))
        DT_SubmittalLogTemplate.Columns.Add(New DataColumn("RevNo", System.Type.GetType("System.Int32")))
        DT_SubmittalLogTemplate.Columns.Add(New DataColumn("Ref1"))
        DT_SubmittalLogTemplate.Columns.Add(New DataColumn("Ref2"))
        DT_SubmittalLogTemplate.Columns.Add(New DataColumn("Ref3"))

        DT_SubmittalLogTemplate.Columns.Add(New DataColumn("Attached"))

        'DT_SubmittalLogTemplate.Columns.Add(New DataColumn("ProDate", System.Type.GetType("System.Date")))
        'DT_SubmittalLogTemplate.Columns.Add(New DataColumn("SubDate", System.Type.GetType("System.Date")))
        'DT_SubmittalLogTemplate.Columns.Add(New DataColumn("DueDate", System.Type.GetType("System.Date")))
        'DT_SubmittalLogTemplate.Columns.Add(New DataColumn("RtnDate", System.Type.GetType("System.Date")))

        DT_SubmittalLogTemplate.Columns.Add(New DataColumn("ProDate", GetType(Date)))
        DT_SubmittalLogTemplate.Columns.Add(New DataColumn("SubDate", GetType(Date)))
        DT_SubmittalLogTemplate.Columns.Add(New DataColumn("DueDate", GetType(Date)))
        DT_SubmittalLogTemplate.Columns.Add(New DataColumn("RtnDate", GetType(Date)))

        'DT_SubmittalLogTemplate.Columns.Add(New DataColumn("ProDate"))
        'DT_SubmittalLogTemplate.Columns.Add(New DataColumn("SubDate"))
        'DT_SubmittalLogTemplate.Columns.Add(New DataColumn("DueDate"))
        'DT_SubmittalLogTemplate.Columns.Add(New DataColumn("RtnDate"))

        DT_SubmittalLogTemplate.Columns.Add(New DataColumn("NoOfDays"))
        DT_SubmittalLogTemplate.Columns.Add(New DataColumn("Status1"))
        Return DT_SubmittalLogTemplate
    End Function

    Public Class csSubmittalLogMain
        Public Str_Flag As String
        Public Str_MenuID As String
        Public Str_FormPrefix As String
        Public dtp_VouDate As Date
        Public int_RevNo As Integer
        Public Str_VouNo As String
        Public Str_Comment As String
        Public str_SalesOrderNo As String
        Public str_ItemCode As String
        Public Int_Slno As Integer

        Public strGUID As String
    End Class

    Public Class csSubmittalLogSub
        Public dt_SubmittalLog As DataTable
    End Class

End Class
