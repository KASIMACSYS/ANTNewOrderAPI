Public Class csBinMovement
    Inherits csSignature
    Public str_SiteID As String
    Public ObjBinMain As New csBinMain
    Public ObjBinSub As New csBinSub
    Public DTBatch As New DataTable

    Public Function DT_StockTrnsTemplate() As DataTable
        DT_StockTrnsTemplate = New DataTable
        DT_StockTrnsTemplate.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
        DT_StockTrnsTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_StockTrnsTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_StockTrnsTemplate.Columns.Add(New DataColumn("FromBin"))
        DT_StockTrnsTemplate.Columns.Add(New DataColumn("ToBin"))
        DT_StockTrnsTemplate.Columns.Add(New DataColumn("Qty", System.Type.GetType("System.Double")))
        DT_StockTrnsTemplate.Columns.Add(New DataColumn("ItemCode"))
        Return DT_StockTrnsTemplate
    End Function
End Class

Public Class csBinMain
    Public str_RefID As String
    Public dtp_RefDate As Date
    Public str_FromWH As String
    Public str_Comment As String
    Public str_CreatedBy As String
    Public dtp_CreatedDate As Date
    Public str_LastUpdatedBy As String
    Public dtp_LastUpdatedDate As Date
    Public int_BusinessPeriodID As Integer
    Public str_ApprovedBy As String
    Public dtp_ApprovedDate As Date
    Public bool_ApprovedStatus As Boolean

    Public int_RevNo As Integer
    Public str_Flag As String
End Class

Public Class csBinSub
    Public dt_BinSub As DataTable
End Class
