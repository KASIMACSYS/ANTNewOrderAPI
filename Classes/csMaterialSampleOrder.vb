'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Public Class csMaterialSampleOrder
    Inherits csSignature
    Public str_SiteID As String
    Public ObjMaterialSampleOrderMain As New csMaterialSampleOrderMain
    Public ObjMaterialSampleOrderSub As New csMaterialSampleOrderSub
    Public objProject As csProjectDetail
    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        objProject = New csProjectDetail
    End Sub
   
    Public Function DT_MaterialSampleOrderTemplate() As DataTable
        DT_MaterialSampleOrderTemplate = New DataTable
        DT_MaterialSampleOrderTemplate.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
        DT_MaterialSampleOrderTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_MaterialSampleOrderTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_MaterialSampleOrderTemplate.Columns.Add(New DataColumn("Unit"))
        DT_MaterialSampleOrderTemplate.Columns.Add(New DataColumn("Qty", System.Type.GetType("System.Double")))
        DT_MaterialSampleOrderTemplate.Columns.Add(New DataColumn("Comment"))
        DT_MaterialSampleOrderTemplate.Columns.Add(New DataColumn("ItemCode"))
        Return DT_MaterialSampleOrderTemplate
    End Function

    Public Function DT_POTDetails() As DataTable
        DT_POTDetails = New DataTable
        DT_POTDetails.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
        DT_POTDetails.Columns.Add(New DataColumn("Section"))
        DT_POTDetails.Columns.Add(New DataColumn("ItemCode"))
        DT_POTDetails.Columns.Add(New DataColumn("Material"))
        DT_POTDetails.Columns.Add(New DataColumn("Qty", System.Type.GetType("System.Double")))
        DT_POTDetails.Columns.Add(New DataColumn("Unit"))
        DT_POTDetails.Columns.Add(New DataColumn("Price", System.Type.GetType("System.Double")))
        DT_POTDetails.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Double")))
        DT_POTDetails.Columns.Add(New DataColumn("Status"))
        Return DT_POTDetails
    End Function
End Class

Public Class csMaterialSampleOrderMain
    Public str_MaterialSampleOrderID As String
    Public str_POTNo As String
    Public str_POTRef As String
    Public str_JobOrderNo As String
    Public dtp_VoucherDate As Date
    Public dtp_IssueDate As Date
    Public dtp_CompletionDate As Date
    Public str_QtnNo As String
    Public int_LedgerID As Integer
    Public str_Alias As String
    Public str_Comment As String
    Public str_Project As String
    Public str_Location As String
    Public str_BrandName As String
    Public str_Contact As String
    Public str_Email As String
    Public str_Item As String
    Public str_Coordinator As String
    Public str_Production As String
    Public str_CreatedBy As String
    Public dtp_CreatedDate As Date
    Public str_LastUpdatedBy As String
    Public dtp_LastUpdatedDate As Date
    Public int_BusinessPeriodID As Integer
    Public str_ApprovedBy As String
    Public dtp_ApprovedDate As Date
    Public bool_ApprovedStatus As Integer
    Public str_CFCCompletionDate As String

    Public str_Prefix As String
    Public int_RevNo As Integer
    Public str_MenuID As String
    Public str_Flag As String

End Class

Public Class csMaterialSampleOrderSub
    Public dt_MaterialSampleOrder As DataTable
End Class


