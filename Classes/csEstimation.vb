Public Class csEstimation
    Inherits csSignature

    Public str_CID As String

    Public objEstMain As New csEstimationMain
    Public objEstSub As New csEstimationSub
    Public objEstVarBOM As New csEstimationVariantBOM
    Public objproject As csProjectDetail

    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        ''If CustomerSetting.Item("Project").ToString = "True" Then
        objproject = New csProjectDetail
        '' End If
    End Sub

    Public Function DT_ESTItemDetailsFGTemplate() As DataTable
        DT_ESTItemDetailsFGTemplate = New DataTable
        DT_ESTItemDetailsFGTemplate.Columns.Add(New DataColumn("SlNo", System.Type.GetType("System.Int32")))
        DT_ESTItemDetailsFGTemplate.Columns.Add(New DataColumn("RefNo", System.Type.GetType("System.Int32")))
        DT_ESTItemDetailsFGTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_ESTItemDetailsFGTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_ESTItemDetailsFGTemplate.Columns.Add(New DataColumn("ItemCode"))
        DT_ESTItemDetailsFGTemplate.Columns.Add(New DataColumn("Unit"))
        DT_ESTItemDetailsFGTemplate.Columns.Add(New DataColumn("BaseUnit", System.Type.GetType("System.Double")))
        DT_ESTItemDetailsFGTemplate.Columns.Add(New DataColumn("VouQty", System.Type.GetType("System.Double")))
        DT_ESTItemDetailsFGTemplate.Columns.Add(New DataColumn("PrimaryQty", System.Type.GetType("System.Double")))
        DT_ESTItemDetailsFGTemplate.Columns.Add(New DataColumn("BOMNo"))
        DT_ESTItemDetailsFGTemplate.Columns.Add(New DataColumn("OrgSlNo", System.Type.GetType("System.Int32")))
        DT_ESTItemDetailsFGTemplate.Columns.Add(New DataColumn("Comment"))
        Return DT_ESTItemDetailsFGTemplate
    End Function

    Public Function DT_BOMParamTemplate() As DataTable
        DT_BOMParamTemplate = New DataTable
        DT_BOMParamTemplate.Columns.Add(New DataColumn("RefNo", System.Type.GetType("System.Int32")))
        DT_BOMParamTemplate.Columns.Add(New DataColumn("BOMNo"))
        DT_BOMParamTemplate.Columns.Add(New DataColumn("Parameter"))
        DT_BOMParamTemplate.Columns.Add(New DataColumn("Value"))
        Return DT_BOMParamTemplate
    End Function

    Public Function DT_ESTItemDetailsRMTemplate() As DataTable
        DT_ESTItemDetailsRMTemplate = New DataTable
        DT_ESTItemDetailsRMTemplate.Columns.Add(New DataColumn("Section"))
        DT_ESTItemDetailsRMTemplate.Columns.Add(New DataColumn("RefNo", System.Type.GetType("System.Int32")))
        DT_ESTItemDetailsRMTemplate.Columns.Add(New DataColumn("Alias1"))
        DT_ESTItemDetailsRMTemplate.Columns.Add(New DataColumn("Alias2"))
        DT_ESTItemDetailsRMTemplate.Columns.Add(New DataColumn("ItemCode"))
        DT_ESTItemDetailsRMTemplate.Columns.Add(New DataColumn("Unit"))
        DT_ESTItemDetailsRMTemplate.Columns.Add(New DataColumn("Qty", System.Type.GetType("System.Double")))
        DT_ESTItemDetailsRMTemplate.Columns.Add(New DataColumn("SellPrice", System.Type.GetType("System.Double")))
        DT_ESTItemDetailsRMTemplate.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Double")))
        DT_ESTItemDetailsRMTemplate.Columns.Add(New DataColumn("Stock", System.Type.GetType("System.Double")))
        DT_ESTItemDetailsRMTemplate.Columns.Add(New DataColumn("Cost", System.Type.GetType("System.Double")))
        Return DT_ESTItemDetailsRMTemplate
    End Function

    Public Function DT_ESTItemDetailsFGTemplateProduction() As DataTable
        DT_ESTItemDetailsFGTemplateProduction = New DataTable
        DT_ESTItemDetailsFGTemplateProduction.Columns.Add(New DataColumn("SlNo", System.Type.GetType("System.Int32")))
        DT_ESTItemDetailsFGTemplateProduction.Columns.Add(New DataColumn("ItemCode"))
        DT_ESTItemDetailsFGTemplateProduction.Columns.Add(New DataColumn("ItemDesc"))
        DT_ESTItemDetailsFGTemplateProduction.Columns.Add(New DataColumn("Qty", System.Type.GetType("System.Double")))
        DT_ESTItemDetailsFGTemplateProduction.Columns.Add(New DataColumn("MfdQty", System.Type.GetType("System.Double")))
        DT_ESTItemDetailsFGTemplateProduction.Columns.Add(New DataColumn("UnitCost", GetType(Double)))
        DT_ESTItemDetailsFGTemplateProduction.Columns.Add(New DataColumn("TotCost", GetType(Double)))
        Return DT_ESTItemDetailsFGTemplateProduction
    End Function


End Class

Public Class csEstimationMain
    Public str_Flag As String
    Public str_FormPrefix As String
    Public str_MenuID As String

    Public int_BusinessPeriodID As Integer
    Public dtp_EstDate As Date
    Public int_RevNo As Integer
    Public str_EstNo As String
    Public str_EstDesc As String
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

    Public str_ProdStage As String
    Public dtp_ProdDate As Date
    Public bit_UpdateInv As Boolean
    Public str_Comment As String
    Public str_InvRef As String = String.Empty
End Class

Public Class csEstimationSub
    Public DT_EstItemDetailsFG As New DataTable
    Public DT_EstItemDetailsFGProd As New DataTable
End Class
Public Class csEstimationVariantBOM
    Public str_EstNo As String
    Public str_EstDesc As String
    Public int_SlNo As Integer

    Public DT_BOMParam As New DataTable
    Public DT_EstItemDetailsRM As New DataTable
End Class
