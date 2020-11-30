'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class csStockAdjustment
    Inherits csSignature
    Public str_SiteID As String
    Public ObjStkAdjMain As New csStkAdjMain
    Public ObjStkAdjSub As New csStkAdjSub
    Public ObjStkAdjCommon As New csStkAdjCommon
    Public objproject As New csProjectDetail
    Public DTBatch As New DataTable
    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        '  If CustomerSetting.Item("useProject").ToString = "True" Then
        'If CustomerSetting.Item("useProject").ToString = "True" Then
        objproject = New csProjectDetail
        'End If
    End Sub

    Public Function DT_StockAdjTemplate() As DataTable
        DT_StockAdjTemplate = New DataTable
        DT_StockAdjTemplate.Columns.Add(New DataColumn("Slno", System.Type.GetType("System.Int32")))
        DT_StockAdjTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_StockAdjTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_StockAdjTemplate.Columns.Add(New DataColumn("Unit"))
        DT_StockAdjTemplate.Columns.Add(New DataColumn("CompanyStock", System.Type.GetType("System.Double")))
        DT_StockAdjTemplate.Columns.Add(New DataColumn("WHStock", System.Type.GetType("System.Double")))
        DT_StockAdjTemplate.Columns.Add(New DataColumn("Excess", System.Type.GetType("System.Double")))
        DT_StockAdjTemplate.Columns.Add(New DataColumn("Shortage", System.Type.GetType("System.Double")))
        DT_StockAdjTemplate.Columns.Add(New DataColumn("WHTotStock", System.Type.GetType("System.Double")))
        DT_StockAdjTemplate.Columns.Add(New DataColumn("Cur_Wac", System.Type.GetType("System.Double")))
        DT_StockAdjTemplate.Columns.Add(New DataColumn("Price", System.Type.GetType("System.Double")))
        DT_StockAdjTemplate.Columns.Add(New DataColumn("Comment"))
        DT_StockAdjTemplate.Columns.Add(New DataColumn("ItemCode"))
        Return DT_StockAdjTemplate
    End Function
End Class

Public Class csStkAdjMain
    Public str_DocumentNo As String
    Public dtp_EntryDate As Date
    Public dtp_DocumentDate As Date
    Public str_Comment As String
    Public str_CreatedBy As String
    Public dtp_CreatedDate As Date
    Public str_LastUpdatedBy As String
    Public dtp_LastUpdatedDate As Date
    Public int_BusinessPeriodID As Integer
    Public int_LanguageCode As Integer
    Public str_WHID As String
    Public int_RevNo As Integer
    Public str_MenuID As String
End Class

Public Class csStkAdjSub
    Public dt_StkAdjSub As DataTable
End Class

Public Class csStkAdjCommon
    Public str_Flag As String
    Public str_stkAdjPrefix As String
End Class

