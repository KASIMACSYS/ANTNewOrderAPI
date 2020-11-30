'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Public Class csItemMaster
    Inherits csSignature
    Public str_SiteID As String
    Public int_BusinessPeriodID As Integer
    Public str_Flag As String
    Public POS As Boolean

    Public str_ProductCode As String
    Public str_ItemDesc1 As String
    Public str_Alias1 As String
    Public str_ItemDesc2 As String
    Public str_Alias2 As String
    Public str_Type As String
    Public str_Method As String
    Public str_Unit1 As String
    Public int_Category1 As Integer
    Public int_Category2 As Integer
    Public int_Category3 As Integer
    Public int_Category4 As Integer
    Public img_Photo() As Byte
    Public str_Desc1 As String
    Public str_Desc2 As String
    Public str_Desc3 As String
    Public str_Desc4 As String
    Public str_Desc5 As String
    Public str_Desc6 As String
    Public str_Desc7 As String
    Public str_Desc8 As String
    Public str_Volume As String
    Public str_VolumeUnit As String
    Public str_Weight As String
    Public str_WeightUnit As String
    Public str_BOM As String
    Public str_POS As String
    Public int_Inventory As Integer
    Public int_CoGS As Integer
    Public bool_BatchTracking As Boolean
    Public bool_BinTracking As Boolean
    Public bool_SerialTracking As Boolean
    Public str_Tax1Code As String
    Public str_Tax2Code As String
    Public str_AttributeOption As String
    Public str_Separator1 As String
    Public str_Separator2 As String
    Public str_Separator3 As String
    Public str_Separator4 As String
    Public str_Separator5 As String
    Public bool_SKUItem As Boolean

    Public DT_XMLData As DataTable
    Public DTItemMaster As DataTable
    Public ObjItemExtraInfo As New csItemExtraInfo
    Public ObjItemLocation As New csItemLocation
    Public ObjItemPrice As New csItemPrice
    Public ObjItemDiscList As New csItemDiscList
    Public objItemOpeningStock As New csItemOpeningStock
    Public ObjParameter As New csItemParameter
    Public ObjItemSecUOM As New csItemSecUOM
    Public ObjDefaulLedger As New csDefaultLedger
    Public ObjItemTax As New csItemTax
    Public ObjItemBarCode As New csItemBarCode

    Public Function DBTemplate() As DataTable
        DTItemMaster = New DataTable
        DTItemMaster.Columns.Add(New DataColumn("ItemCode"))
        DTItemMaster.Columns.Add(New DataColumn("ItemDesc1"))
        DTItemMaster.Columns.Add(New DataColumn("Alias1"))
        DTItemMaster.Columns.Add(New DataColumn("ItemDesc2"))
        DTItemMaster.Columns.Add(New DataColumn("Alias2"))
        DTItemMaster.Columns.Add(New DataColumn("BarCodeNo"))
        DTItemMaster.Columns.Add(New DataColumn("InActive", System.Type.GetType("System.Boolean")))
        DTItemMaster.Columns.Add(New DataColumn("Approved", System.Type.GetType("System.Boolean")))
        DTItemMaster.Columns.Add(New DataColumn("MinQty", System.Type.GetType("System.Double")))
        DTItemMaster.Columns.Add(New DataColumn("MaxQty", System.Type.GetType("System.Double")))
        DTItemMaster.Columns.Add(New DataColumn("MinSellPrice", System.Type.GetType("System.Double")))
        DTItemMaster.Columns.Add(New DataColumn("MaxPurPrice", System.Type.GetType("System.Double")))
        DTItemMaster.Columns.Add(New DataColumn("Edit"))

        DTItemMaster.Columns("MinQty").DefaultValue = 0
        DTItemMaster.Columns("MaxQty").DefaultValue = 0
        DTItemMaster.Columns("MinSellPrice").DefaultValue = 0
        DTItemMaster.Columns("MaxPurPrice").DefaultValue = 0
        DTItemMaster.Columns("InActive").DefaultValue = False

        Return DTItemMaster
    End Function
End Class

Public Class csItemExtraInfo
    Public str_Cat1 As String
    Public str_Cat2 As String
    Public str_Cat3 As String
    Public str_Cat4 As String
   
End Class

Public Class csItemTax
    Public str_VAT_Type As String
    Public dbl_VAT_Percent As Double
    Public dbl_VAT_Amt As Double
    Public str_Tax As String
    'Public str_Tax2Code As String
    Public dbl_TaxSalesPercent As Double
    Public dbl_TaxPurchasePercent As Double
    'Public dbl_Tax2SalesPercent As Double
    'Public dbl_Tax2PurchasePercent As Double
End Class

Public Class csItemLocation
    Public dt_StockInfo As New DataTable
End Class

Public Class csItemPrice
    Public dt_PriceList As New DataTable
End Class
Public Class csItemDiscList
    Public dt_ItemDiscList As New DataTable
End Class

Public Class csItemBarCode
    Public dt_BarCode As New DataTable
End Class

Public Class csItemOpeningStock
    Public str_BusinessPeriodID As String
    Public dbl_OpenStock As Double
    Public dbl_OpenWAC As Double
End Class

Public Class csItemParameter
    Public dt_Parameter As New DataTable
End Class
Public Class csItemSecUOM
    Public dt_ItemSecUOM As New DataTable
End Class
Public Class csAttibute
    Public str_AttributeName1 As String
    Public str_AttributeName2 As String
    Public str_AttributeName As String
    Public str_AttributeCode As String
    Public bool_AttributeOption As Boolean
    Public dt_AttributeName As DataTable
    Public dt_AttributeOption As DataTable
    Public objItemMaster As New csItemMaster

    Public Function DBTemplate() As DataTable
        Dim DT_Template As New DataTable
        DT_Template.Columns.Add(New DataColumn("AttributeCode"))
        DT_Template.Columns.Add(New DataColumn("Code"))
        DT_Template.Columns.Add(New DataColumn("Options1"))
        DT_Template.Columns.Add(New DataColumn("Options2"))
        Return DT_Template
    End Function
End Class




