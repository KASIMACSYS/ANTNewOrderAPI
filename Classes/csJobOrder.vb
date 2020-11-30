'======================================================================================
'$Author: Amjath $
'$Rev: 3537 $
'$Date: 2018-05-18 15:47:58 +0530 (Fri, 18 May 2018) $ 
'======================================================================================
Public Class csJobOrder
    Inherits csSignature

    Public str_CID As String

    Public objJOMain As New csJobOrderMain
    Public objJOSub As New csJobOrderSub
    Public objJOVarBOM As New csJobOrderVariantBOM
    Public objproject As csProjectDetail

    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        ''If CustomerSetting.Item("Project").ToString = "True" Then
        objproject = New csProjectDetail
        '' End If
    End Sub

    Public Function DT_JOItemDetailsFGTemplate() As DataTable
        DT_JOItemDetailsFGTemplate = New DataTable
        DT_JOItemDetailsFGTemplate.Columns.Add(New DataColumn("SlNo", System.Type.GetType("System.Int32")))
        DT_JOItemDetailsFGTemplate.Columns.Add(New DataColumn("RefNo", System.Type.GetType("System.Int32")))
        DT_JOItemDetailsFGTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_JOItemDetailsFGTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_JOItemDetailsFGTemplate.Columns.Add(New DataColumn("ItemCode"))
        DT_JOItemDetailsFGTemplate.Columns.Add(New DataColumn("Unit"))
        DT_JOItemDetailsFGTemplate.Columns.Add(New DataColumn("BaseUnit", System.Type.GetType("System.Double")))
        DT_JOItemDetailsFGTemplate.Columns.Add(New DataColumn("VouQty", System.Type.GetType("System.Double")))
        DT_JOItemDetailsFGTemplate.Columns.Add(New DataColumn("PrimaryQty", System.Type.GetType("System.Double")))
        DT_JOItemDetailsFGTemplate.Columns.Add(New DataColumn("BOMNo"))
        DT_JOItemDetailsFGTemplate.Columns.Add(New DataColumn("OrgSlNo", System.Type.GetType("System.Int32")))
        DT_JOItemDetailsFGTemplate.Columns.Add(New DataColumn("Comment"))
        Return DT_JOItemDetailsFGTemplate
    End Function

    Public Function DT_BOMParamTemplate() As DataTable
        DT_BOMParamTemplate = New DataTable
        DT_BOMParamTemplate.Columns.Add(New DataColumn("RefNo", System.Type.GetType("System.Int32")))
        DT_BOMParamTemplate.Columns.Add(New DataColumn("BOMNo"))
        DT_BOMParamTemplate.Columns.Add(New DataColumn("Parameter"))
        DT_BOMParamTemplate.Columns.Add(New DataColumn("Value"))
        Return DT_BOMParamTemplate
    End Function

    Public Function DT_JOItemDetailsRMTemplate() As DataTable
        DT_JOItemDetailsRMTemplate = New DataTable
        DT_JOItemDetailsRMTemplate.Columns.Add(New DataColumn("RefNo", System.Type.GetType("System.Int32")))
        DT_JOItemDetailsRMTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_JOItemDetailsRMTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_JOItemDetailsRMTemplate.Columns.Add(New DataColumn("ItemCode"))
        DT_JOItemDetailsRMTemplate.Columns.Add(New DataColumn("Unit"))
        DT_JOItemDetailsRMTemplate.Columns.Add(New DataColumn("Qty", System.Type.GetType("System.Double")))
        DT_JOItemDetailsRMTemplate.Columns.Add(New DataColumn("Cost", System.Type.GetType("System.Double")))
        DT_JOItemDetailsRMTemplate.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Double")))
        DT_JOItemDetailsRMTemplate.Columns.Add(New DataColumn("Stock", System.Type.GetType("System.Double")))
        Return DT_JOItemDetailsRMTemplate
    End Function

    Public Function DT_JOItemDetailsFGTemplateProduction() As DataTable
        DT_JOItemDetailsFGTemplateProduction = New DataTable
        DT_JOItemDetailsFGTemplateProduction.Columns.Add(New DataColumn("SlNo", System.Type.GetType("System.Int32")))
        DT_JOItemDetailsFGTemplateProduction.Columns.Add(New DataColumn("ItemCode"))
        DT_JOItemDetailsFGTemplateProduction.Columns.Add(New DataColumn("ItemDesc"))
        DT_JOItemDetailsFGTemplateProduction.Columns.Add(New DataColumn("Qty", System.Type.GetType("System.Double")))
        DT_JOItemDetailsFGTemplateProduction.Columns.Add(New DataColumn("MfdQty", System.Type.GetType("System.Double")))
        DT_JOItemDetailsFGTemplateProduction.Columns.Add(New DataColumn("UnitCost", GetType(Double)))
        DT_JOItemDetailsFGTemplateProduction.Columns.Add(New DataColumn("TotCost", GetType(Double)))
        Return DT_JOItemDetailsFGTemplateProduction
    End Function

 
End Class

Public Class csJobOrderMain
    Public str_Flag As String
    Public str_FormPrefix As String
    Public str_MenuID As String

    Public int_BusinessPeriodID As Integer
    Public dtp_JODate As Date
    Public str_SONo As String
    Public int_RevNo As Integer
    Public str_JONo As String
    Public str_JODesc As String
    Public int_LedgerID As Integer
    Public int_SalesManID As Int32
    Public str_ProdUnitName As String

    Public dtp_EstEndDate As Date
    Public dtp_ActEndDate As Date
    Public dbl_ManDays As Double
    Public dbl_EstCost As Double
    Public dbl_ActCost As Double
    Public dbl_EstMatCost As Double
    Public dbl_ActMatCost As Double
    Public str_Status As String
    Public str_LpoNo As String
    Public str_FG As String

    Public str_ProdStage As String
    Public dtp_ProdDate As Date
    Public bit_UpdateInv As Boolean
    Public str_Comment As String
End Class

Public Class csJobOrderSub
    Public DT_JOItemDetailsFG As New DataTable
    Public DT_JOItemDetailsFGProd As New DataTable
End Class
Public Class csJobOrderVariantBOM
    Public str_JONo As String
    Public str_JODesc As String
    Public int_SlNo As Integer

    Public DT_BOMParam As New DataTable
    Public DT_JOItemDetailsRM As New DataTable
End Class