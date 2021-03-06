﻿Public Class csBarCode
    Public Sub ItemDetails(ByVal DTItem As DataTable, ByRef ItmDetails As Hashtable, ByVal _Stock As Double, ByVal _Cost As Double)
        ItmDetails = New Hashtable
        ItmDetails.Add("ItemCode", DTItem(0)("ItemCode").ToString)
        ItmDetails.Add("ItemDesc1", DTItem(0)("ItemDesc1").ToString)
        ItmDetails.Add("Alias1", DTItem(0)("Alias1").ToString)
        ItmDetails.Add("ItemDesc2", DTItem(0)("ItemDesc2").ToString)
        ItmDetails.Add("Alias2", DTItem(0)("Alias2").ToString)
        ItmDetails.Add("Unit", DTItem(0)("Unit").ToString)
        ItmDetails.Add("Cost", _Cost)
        ItmDetails.Add("Type", DTItem(0)("Type").ToString)
        ItmDetails.Add("Desc1", DTItem(0)("Desc1").ToString)
        ItmDetails.Add("Desc2", DTItem(0)("Desc2").ToString)
        ItmDetails.Add("Desc3", DTItem(0)("Desc3").ToString)
        ItmDetails.Add("Desc4", DTItem(0)("Desc4").ToString)
        ItmDetails.Add("Desc5", DTItem(0)("Desc5").ToString)
        ItmDetails.Add("Desc6", DTItem(0)("Desc6").ToString)
        ItmDetails.Add("Desc7", DTItem(0)("Desc7").ToString)
        ItmDetails.Add("Desc8", DTItem(0)("Desc8").ToString)
        ItmDetails.Add("CostType", "")
        ItmDetails.Add("BOMNo", DTItem(0)("BOMNo").ToString)
        'ItmDetails.Add("Unit2", DTItem(0)("Unit2").ToString)
        'ItmDetails.Add("Ratio", DTItem(0)("Ratio").ToString)
        ItmDetails.Add("MaxPurPrice", DTItem(0)("MaxPurPrice").ToString)
        ItmDetails.Add("MinSellPrice", DTItem(0)("MinSellPrice").ToString)
        ItmDetails.Add("SellPrice", _Cost)
        ItmDetails.Add("EnquiryPrice", _Cost)
        ItmDetails.Add("Stock", _Stock)
        ItmDetails.Add("BarCodeNo", DTItem(0)("BarCode"))
        ItmDetails.Add("Inventory", DTItem(0)("Inventory"))
        ItmDetails.Add("Batch", DTItem(0)("Batch"))
        ItmDetails.Add("Bin", DTItem(0)("Bin"))
        ItmDetails.Add("Serial", DTItem(0)("Serial"))
        ItmDetails.Add("Tax", DTItem(0)("Tax"))
        ItmDetails.Add("WHStock", _Stock)
        ItmDetails.Add("AvlStock", _Stock)
    End Sub

End Class
