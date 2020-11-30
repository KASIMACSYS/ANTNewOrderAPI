'======================================================================================
'$Author: Kasim $
'$Rev: 510 $
'$Date: 2016-07-27 19:50:21 +0530 (Wed, 27 July 2016) $ 
'======================================================================================
Public Class csAirwayBill
    Inherits csSignature

    Public str_SiteID As String

    Public objAirwayBillMain As New AirwayBillMain
    Public DT_AirwayBillSub As New DataTable

    Public Function DT_AirwayBillTemplate() As DataTable
        DT_AirwayBillTemplate = New DataTable
        DT_AirwayBillTemplate.Columns.Add(New DataColumn("SlNo", System.Type.GetType("System.Int32")))
        DT_AirwayBillTemplate.Columns.Add(New DataColumn("NoOfPCS", System.Type.GetType("System.Double")))
        DT_AirwayBillTemplate.Columns.Add(New DataColumn("GrossWgt", System.Type.GetType("System.Decimal")))
        DT_AirwayBillTemplate.Columns.Add(New DataColumn("RateClass", System.Type.GetType("System.Decimal")))
        DT_AirwayBillTemplate.Columns.Add(New DataColumn("ChagWeight", System.Type.GetType("System.Decimal")))
        DT_AirwayBillTemplate.Columns.Add(New DataColumn("RateCharge", System.Type.GetType("System.Decimal")))
        DT_AirwayBillTemplate.Columns.Add(New DataColumn("Total", System.Type.GetType("System.Decimal")))
        DT_AirwayBillTemplate.Columns.Add(New DataColumn("QualityOfGoods"))
        Return DT_AirwayBillTemplate
    End Function
End Class

Public Class AirwayBillMain
    Public str_Flag As String
    Public str_FormPrefix As String
    Public str_MenuID As String
    Public int_BusinessPeriodID As Integer
    Public int_RevNo As Integer
    Public dtp_Date As Date
    Public str_RefNo As String
    Public str_BOMDesc As String
    Public str_Text1 As String
    Public str_Text2 As String
    Public str_Text3 As String
    Public str_ShipperName As String
    Public str_IssuedBy As String
    Public str_ConsignName As String
    Public str_IssueAgent As String
    Public str_AccInfo As String
    Public str_AirportDept As String
    Public str_AirportTo As String
    Public str_ByFirstCarrier As String
    Public str_Currency As String
    Public str_ChqCode As String
    Public str_WTVALPPD As String
    Public str_WTVALCOLL As String
    Public str_OtherPPD As String
    Public str_OtherCOLL As String
    Public str_ValueForCarriage As String
    Public str_ValueForCustoms As String
    Public str_AirportofDest As String
    Public str_HandlingInfo As String
    Public str_TotalPCS As String
    Public str_TotalGrossWeight As String
    Public str_Total As String
    Public str_WCPrepaid As String
    Public str_WCCollect As String
    Public str_ValueChargePrepaid As String
    Public str_ValueChargeCollect As String
    Public str_TaxPrepaid As String
    Public str_TaxCollect As String

    Public str_DueAgentPrepaid As String
    Public str_DueAgentCollect As String
    Public str_DueCarrierPrepaid As String
    Public str_DueCarrierCollect As String
    Public str_OtherCharge As String
    Public str_TotalPrepaid As String
    Public str_TotalCollect As String
    Public str_ExecDate As String
    Public str_ExecPlace As String
    Public str_Extra1 As String
End Class