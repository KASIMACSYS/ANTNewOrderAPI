'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Public Class csStockTransfer
    Inherits csSignature
    Public str_SiteID As String
    Public ObjStkTrnsMain As New csStkTrnsMain
    Public ObjStkTrnsSub As New csStkTrnsSub
    Public DTBatch As New DataTable

    Public Function DT_StockTrnsTemplate() As DataTable
        DT_StockTrnsTemplate = New DataTable
        DT_StockTrnsTemplate.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
        DT_StockTrnsTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_StockTrnsTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_StockTrnsTemplate.Columns.Add(New DataColumn("Unit"))
        DT_StockTrnsTemplate.Columns.Add(New DataColumn("TotQty", System.Type.GetType("System.Double")))
        DT_StockTrnsTemplate.Columns.Add(New DataColumn("PrimaryQty", System.Type.GetType("System.Double")))
        DT_StockTrnsTemplate.Columns.Add(New DataColumn("Comment"))
        DT_StockTrnsTemplate.Columns.Add(New DataColumn("ItemCode"))
        Return DT_StockTrnsTemplate
    End Function
End Class

Public Class csStkTrnsMain
    Public str_TransferID As String
    Public dtp_TransferDate As Date
    Public str_FromWH As String
    Public str_ToWH As String
    Public str_Comment As String
    Public str_CreatedBy As String
    Public dtp_CreatedDate As Date
    Public str_LastUpdatedBy As String
    Public dtp_LastUpdatedDate As Date
    Public int_BusinessPeriodID As Integer
    Public str_ApprovedBy As String
    Public dtp_ApprovedDate As Date
    Public bool_ApprovedStatus As Integer
    Public int_LanguageCode As Integer
    Public str_Prefix As String
    Public int_RevNo As Integer
    Public str_MenuID As String
    Public str_Flag As String
    Public str_StockTransferStatus As String
End Class

Public Class csStkTrnsSub
    Public dt_StkTrnsSub As DataTable
End Class


