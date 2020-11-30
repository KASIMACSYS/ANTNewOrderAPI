
Public Class csPOT
    Inherits csSignature
    Public str_CID As String
    Public int_BusinessPeriodID As Integer
    Public ObjPOTMain As New csPOTMain
    Public ObjPOTSub As New csPOTSub
    Public objproject As New csProjectDetail

    Public Function DT_POTTemplate() As DataTable
        DT_POTTemplate = New DataTable
        DT_POTTemplate.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
        DT_POTTemplate.Columns.Add(New DataColumn("ItemCode"))
        DT_POTTemplate.Columns.Add(New DataColumn("ItemDesc"))
        DT_POTTemplate.Columns.Add(New DataColumn("Status"))
        DT_POTTemplate.Columns.Add(New DataColumn("BOQ"))
        DT_POTTemplate.Columns.Add(New DataColumn("DrawingNo"))
        DT_POTTemplate.Columns.Add(New DataColumn("Area"))
        DT_POTTemplate.Columns.Add(New DataColumn("Unit"))
        DT_POTTemplate.Columns.Add(New DataColumn("VouQty", System.Type.GetType("System.Double")))
        'DT_POTTemplate.Columns.Add(New DataColumn("Price", System.Type.GetType("System.Double")))
        'DT_POTTemplate.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Double")))
        Return DT_POTTemplate
    End Function

    Public Class csPOTMain
        Public Str_Flag As String
        Public Str_MenuID As String
        Public Str_FormPrefix As String
        Public dtp_VouDate As Date
        Public int_RevNo As Integer
        Public Str_VouNo As String
        Public Str_Comment As String
        Public str_JONo As String
        Public dtp_StartDate As Date
        Public dtp_EndDate As Date
        Public str_ItemCode As String
        Public str_ItemDesc As String
        Public str_Status As String
        Public str_BOQ As String
        Public str_SONo As String
        Public str_DrawingNo As String
        Public str_Area As String
        Public str_Unit As String
        Public str_Qty As String
        Public str_Price As String
        Public str_Amount As String
        Public str_POTNo As String
        Public str_POTDesc As String

        Public str_ApprovedBy As String
        Public dtp_ApprovedDate As Date
        Public str_IssuedBy As String
        Public dtp_IssuedDate As Date
        Public str_ReceivedBy As String
        Public str_FactoryTeam As String
        Public int_StatusCancel As Integer
        Public int_RevisionNo As Integer
        Public int_OrgSlno As Integer
    End Class

    Public Class csPOTSub
        Public DT_POT As DataTable
        Public dt_Attachment As DataTable
        Public dt_Section As DataTable
        Public Function DT_SectionTemplate() As DataTable
            DT_SectionTemplate = New DataTable
            DT_SectionTemplate.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
            DT_SectionTemplate.Columns.Add(New DataColumn("Section"))
            DT_SectionTemplate.Columns.Add(New DataColumn("Date_"))
            DT_SectionTemplate.Columns.Add(New DataColumn("Comment"))
            Return DT_SectionTemplate
        End Function
        Public Function DT_AttachmentTemplate() As DataTable
            DT_AttachmentTemplate = New DataTable
            DT_AttachmentTemplate.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
            DT_AttachmentTemplate.Columns.Add(New DataColumn("AttachmentType"))
            Return DT_AttachmentTemplate
        End Function
    End Class

End Class
