'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================



Imports Classes
Imports System.Data.SqlClient

Public Class DAL_ItemMaster

#Region "Variable Declaration"
    Dim objcsItemmastr As New csItemMaster
    Dim objcsAttribute As New csAttibute
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General
#End Region

#Region "Get Structure"
    '======================================================================================
    ' Function Name : Get_Structure
    ' Description   : 
    '======================================================================================
    Public Sub Get_Structure(ByRef obj_Struct As csItemMaster, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetItemMaster]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj_Struct.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj_Struct.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj_Struct.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@ProductCode", obj_Struct.str_ProductCode)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDesc1", obj_Struct.str_ItemDesc1)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDesc2", obj_Struct.str_ItemDesc2)
            BaseConn.cmd.Parameters.AddWithValue("@Count", 0)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            obj_Struct.str_ItemDesc1 = ds.Tables(0).Rows(0)("ItemDesc1").ToString()
            obj_Struct.str_Alias1 = ds.Tables(0).Rows(0)("Alias1").ToString()
            obj_Struct.str_ItemDesc2 = ds.Tables(0).Rows(0)("ItemDesc2").ToString()
            obj_Struct.str_Alias2 = ds.Tables(0).Rows(0)("Alias2").ToString()
            obj_Struct.str_Type = ds.Tables(0).Rows(0)("Type").ToString()
            obj_Struct.str_Method = ds.Tables(0).Rows(0)("Method").ToString()
            obj_Struct.str_Unit1 = ds.Tables(0).Rows(0)("Unit1").ToString()
            obj_Struct.str_Separator1 = ds.Tables(0).Rows(0)("Separator1").ToString()
            obj_Struct.str_Separator2 = ds.Tables(0).Rows(0)("Separator2").ToString()
            obj_Struct.str_Separator3 = ds.Tables(0).Rows(0)("Separator3").ToString()
            obj_Struct.str_Separator4 = ds.Tables(0).Rows(0)("Separator4").ToString()
            obj_Struct.str_Separator5 = ds.Tables(0).Rows(0)("Separator5").ToString()
            obj_Struct.bool_SKUItem = ds.Tables(0).Rows(0)("SKUItem")

            obj_Struct.str_Volume = ds.Tables(0).Rows(0)("Volume").ToString()
            obj_Struct.str_VolumeUnit = ds.Tables(0).Rows(0)("VolumeUnit").ToString()
            obj_Struct.str_Weight = ds.Tables(0).Rows(0)("Weight").ToString()
            obj_Struct.str_WeightUnit = ds.Tables(0).Rows(0)("WeightUnit").ToString()

            obj_Struct.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
            obj_Struct.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
            obj_Struct.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
            obj_Struct.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
            obj_Struct.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
            obj_Struct.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
            obj_Struct.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
            obj_Struct.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()

            obj_Struct.int_Category1 = ds.Tables(0).Rows(0)("Category1").ToString()
            obj_Struct.int_Category2 = ds.Tables(0).Rows(0)("Category2").ToString()
            obj_Struct.int_Category3 = ds.Tables(0).Rows(0)("Category3").ToString()
            obj_Struct.int_Category4 = ds.Tables(0).Rows(0)("Category4").ToString()

            obj_Struct.bool_BatchTracking = ds.Tables(0).Rows(0)("BatchTracking")
            obj_Struct.bool_SerialTracking = ds.Tables(0).Rows(0)("SerialTracking")
            obj_Struct.bool_BinTracking = ds.Tables(0).Rows(0)("BinTracking")

            obj_Struct.int_Inventory = ds.Tables(0).Rows(0)("Inventory")
            obj_Struct.int_CoGS = ds.Tables(0).Rows(0)("CoGS")
            obj_Struct.str_Tax1Code = ds.Tables(0).Rows(0)("Tax")
            obj_Struct.ObjItemTax.str_Tax = ds.Tables(0).Rows(0)("Tax")

            Dim str As String = ds.Tables(0).Rows(0)("Photo").ToString()
            If str.Length > 0 Then
                obj_Struct.img_Photo = CType(ds.Tables(0).Rows(0)("Photo"), Byte())
            Else
                obj_Struct.img_Photo = Nothing
            End If

            obj_Struct.DTItemMaster = ds.Tables(1)

            If ds.Tables(2).Rows.Count > 0 Then
                obj_Struct.ObjItemExtraInfo.str_Cat1 = ds.Tables(2).Rows(0)("Cat1").ToString()
                obj_Struct.ObjItemExtraInfo.str_Cat2 = ds.Tables(2).Rows(0)("Cat2").ToString()
                obj_Struct.ObjItemExtraInfo.str_Cat3 = ds.Tables(2).Rows(0)("Cat3").ToString()
                obj_Struct.ObjItemExtraInfo.str_Cat4 = ds.Tables(2).Rows(0)("Cat4").ToString()
            End If


            obj_Struct.ObjItemPrice.dt_PriceList = ds.Tables(3)

            'obj_Struct.ObjItemLocation.dt_StockInfo = ds.Tables(4)

            'If ds.Tables(5).Rows.Count > 0 Then
            '    obj_Struct.ObjItemTax.str_VAT_Type = ds.Tables(5).Rows(0)("TaxType").ToString()
            '    obj_Struct.ObjItemTax.dbl_VAT_Percent = ds.Tables(5).Rows(0)("TaxPercent").ToString()
            '    obj_Struct.ObjItemTax.dbl_VAT_Amt = ds.Tables(5).Rows(0)("TaxAmount").ToString()
            'Else
            '    obj_Struct.ObjItemTax.str_VAT_Type = ""
            '    obj_Struct.ObjItemTax.dbl_VAT_Percent = 0
            '    obj_Struct.ObjItemTax.dbl_VAT_Amt = 0
            'End If
            obj_Struct.ObjItemBarCode.dt_BarCode = ds.Tables(12)

            obj_Struct.ObjItemDiscList.dt_ItemDiscList = ds.Tables(4)
            obj_Struct.ObjParameter.dt_Parameter = ds.Tables(5)
            obj_Struct.ObjItemSecUOM.dt_ItemSecUOM = ds.Tables(6)
            'If ds.Tables(7).Rows.Count > 0 Then
            '    obj_Struct.ObjItemTax.str_Tax = ds.Tables(7).Rows(0)("TaxCode").ToString()
            '    obj_Struct.ObjItemTax.dbl_TaxSalesPercent = ds.Tables(7).Rows(0)("SalesTaxPercentage").ToString()
            '    obj_Struct.ObjItemTax.dbl_TaxPurchasePercent = ds.Tables(7).Rows(0)("PurchaseTaxPercentage").ToString()
            'Else
            '    obj_Struct.ObjItemTax.str_Tax = ""
            '    obj_Struct.ObjItemTax.dbl_TaxSalesPercent = 0
            '    obj_Struct.ObjItemTax.dbl_TaxPurchasePercent = 0
            'End If
            'If ds.Tables(8).Rows.Count > 0 Then
            '    obj_Struct.ObjItemTax.str_Tax = ds.Tables(8).Rows(0)("TaxCode").ToString()
            '    obj_Struct.ObjItemTax.dbl_TaxSalesPercent = ds.Tables(8).Rows(0)("SalesTaxPercentage").ToString()
            '    obj_Struct.ObjItemTax.dbl_TaxPurchasePercent = ds.Tables(8).Rows(0)("PurchaseTaxPercentage").ToString()
            '    'Else
            '    '    obj_Struct.ObjItemTax.str_Tax2Code = ""
            '    '    obj_Struct.ObjItemTax.dbl_Tax2SalesPercent = 0
            '    '    obj_Struct.ObjItemTax.dbl_Tax2PurchasePercent = 0
            'End If


            obj_Struct.str_BOM = ds.Tables(0).Rows(0)("BOMNo").ToString()
            If ds.Tables(7).Rows.Count > 0 Then
                obj_Struct.str_POS = ds.Tables(7).Rows(0)("POSCategory").ToString()
            End If

            If ds.Tables(8).Rows.Count > 0 Then
                obj_Struct.str_AttributeOption = ds.Tables(8).Rows(0)("AttributeOptions")
            End If
            
            obj_Struct.DT_XMLData = ds.Tables(9)

            obj_Struct.DTItemMaster = ds.Tables(10)

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    '======================================================================================
    ' Function Name : Get_Structure_UoM
    ' Description   : 
    '======================================================================================
    Public Sub Get_Structure_UoM(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _BusinessPeriodID As Integer, ByVal Str_ItemCode As String, ByVal Str_ItemDesc As String, ByVal Str_Flag As String, ByRef UOMCount As Integer)

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetItemMaster]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@ProductCode", Str_ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDesc1", Str_ItemDesc)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDesc2", Str_ItemDesc)
            BaseConn.cmd.Parameters.Add("@Count", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            UOMCount = BaseConn.cmd.Parameters("@Count").Value.ToString
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub
#End Region

#Region "Update ItemMaster"
    '======================================================================================
    ' Function Name : Insert_ItemMaster
    ' Description   : 
    '======================================================================================
    ''' Public Function Insert_ItemMaster(ByVal objFromArg As csItemMaster, ByRef ItemCode As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
    Public Sub Insert_ItemMaster(ByVal objFromArg As csItemMaster, ByRef ItemCode As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String,
        ByRef ErrNo As Integer, ByRef _ErrString As String)

        _ErrString = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("ItemmasterUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", objFromArg.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", objFromArg.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", objFromArg.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@POS", objFromArg.str_POS)
            BaseConn.cmd.Parameters.AddWithValue("@ProductCode", objFromArg.str_ProductCode)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDesc1", objFromArg.str_ItemDesc1)
            BaseConn.cmd.Parameters.AddWithValue("@Alias1", objFromArg.str_Alias1)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDesc2", objFromArg.str_ItemDesc2)
            BaseConn.cmd.Parameters.AddWithValue("@Alias2", objFromArg.str_Alias2)
            BaseConn.cmd.Parameters.AddWithValue("@Type", objFromArg.str_Type)
            BaseConn.cmd.Parameters.AddWithValue("@Method", objFromArg.str_Method)
            BaseConn.cmd.Parameters.AddWithValue("@Unit1", objFromArg.str_Unit1)
            BaseConn.cmd.Parameters.AddWithValue("@Category1", objFromArg.int_Category1)
            BaseConn.cmd.Parameters.AddWithValue("@Category2", objFromArg.int_Category2)
            BaseConn.cmd.Parameters.AddWithValue("@Category3", objFromArg.int_Category3)
            BaseConn.cmd.Parameters.AddWithValue("@Category4", objFromArg.int_Category4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc1", objFromArg.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", objFromArg.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", objFromArg.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", objFromArg.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", objFromArg.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", objFromArg.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", objFromArg.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", objFromArg.str_Desc8)

            If objFromArg.img_Photo Is Nothing Then
                Dim photoParam As New SqlParameter("@Photo", SqlDbType.Image)
                photoParam.Value = DBNull.Value
                BaseConn.cmd.Parameters.Add(photoParam)
            Else
                BaseConn.cmd.Parameters.AddWithValue("@Photo", objFromArg.img_Photo) 'PHOTO
            End If

            BaseConn.cmd.Parameters.AddWithValue("@Volume", objFromArg.str_Volume)
            BaseConn.cmd.Parameters.AddWithValue("@VolumeUnit", objFromArg.str_VolumeUnit)
            BaseConn.cmd.Parameters.AddWithValue("@Weight", objFromArg.str_Weight)
            BaseConn.cmd.Parameters.AddWithValue("@WeightUnit", objFromArg.str_WeightUnit)

            BaseConn.cmd.Parameters.AddWithValue("@BOM", objFromArg.str_BOM)
            BaseConn.cmd.Parameters.AddWithValue("@Inventory", objFromArg.int_Inventory)
            BaseConn.cmd.Parameters.AddWithValue("@CoGS", objFromArg.int_CoGS)
            BaseConn.cmd.Parameters.AddWithValue("@BatchTracking", objFromArg.bool_BatchTracking)
            BaseConn.cmd.Parameters.AddWithValue("@SerialTracking", objFromArg.bool_SerialTracking)
            BaseConn.cmd.Parameters.AddWithValue("@BinTracking", objFromArg.bool_BinTracking)
            BaseConn.cmd.Parameters.AddWithValue("@Tax", objFromArg.ObjItemTax.str_Tax)
            'BaseConn.cmd.Parameters.AddWithValue("@Tax2Code", objFromArg.ObjItemTax.str_Tax2Code)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", objFromArg.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", objFromArg.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", objFromArg.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", objFromArg.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", objFromArg.ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@Separator1", objFromArg.str_Separator1)
            BaseConn.cmd.Parameters.AddWithValue("@Separator2", objFromArg.str_Separator2)
            BaseConn.cmd.Parameters.AddWithValue("@Separator3", objFromArg.str_Separator3)
            BaseConn.cmd.Parameters.AddWithValue("@Separator4", objFromArg.str_Separator4)
            BaseConn.cmd.Parameters.AddWithValue("@Separator5", objFromArg.str_Separator5)
            BaseConn.cmd.Parameters.AddWithValue("@SKUItem", objFromArg.bool_SKUItem)

            BaseConn.cmd.Parameters.AddWithValue("@AttributeOptions", objFromArg.str_AttributeOption)

            BaseConn.cmd.Parameters.AddWithValue("@ItemMasterDT", objFromArg.DTItemMaster) 'ItemMaster
            BaseConn.cmd.Parameters.AddWithValue("@ItemParameterDT", objFromArg.ObjParameter.dt_Parameter) 'BOM Parameter
            BaseConn.cmd.Parameters.AddWithValue("@LocationDT", objFromArg.ObjItemLocation.dt_StockInfo) 'ItemLocation
            BaseConn.cmd.Parameters.AddWithValue("@ItemPriceDT", objFromArg.ObjItemPrice.dt_PriceList) 'ItemPrice
            BaseConn.cmd.Parameters.AddWithValue("@ItemSecUOMDT", objFromArg.ObjItemSecUOM.dt_ItemSecUOM) 'ItemSecUOM
            BaseConn.cmd.Parameters.AddWithValue("@ItemDiscListDT", objFromArg.ObjItemDiscList.dt_ItemDiscList) 'ItemDiscList
            BaseConn.cmd.Parameters.AddWithValue("@ItemBarCodeDT", objFromArg.ObjItemBarCode.dt_BarCode) 'ItemBarCode

            BaseConn.cmd.Parameters.Add("@ProductCodeOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            ItemCode = BaseConn.cmd.Parameters("@ProductCodeOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(objFromArg.str_SiteID)
            ObjDalGeneral.Elog_Insert(objFromArg.str_SiteID, _StrDBPath, _StrDBPwd, objFromArg.int_BusinessPeriodID, objFromArg.str_CreatedBy, objFromArg.dtp_CreatedDate, "", "ItemMaster", Err.Number, "Error in " & objFromArg.str_Flag & " : " & objFromArg.str_ProductCode & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetItemPrice(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _StrSiteID As String, ByVal _ItemCode As String, ByVal _PriceType As String, _
                            ByRef ItemMarkUpByPercentage As Double, ByRef ItemMarkUpByPrice As Double)

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetItemPrice]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _StrSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@PriceType", _PriceType)
            BaseConn.cmd.Parameters.Add("@ItemMarkUpByPercentage", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ItemMarkUpByPrice", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            ItemMarkUpByPercentage = BaseConn.cmd.Parameters("@ItemMarkUpByPercentage").Value
            ItemMarkUpByPrice = BaseConn.cmd.Parameters("@ItemMarkUpByPrice").Value
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub
#End Region

    Public Function GetItemsDetailsFromGivenItems(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal GivenItems As DataTable, ByVal _Flag As String, ByRef ErrNo As Integer) As DataTable
        Dim _ErrString As String = ""
        ErrNo = 0
        GetItemsDetailsFromGivenItems = New DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetItemsDetailsFromGivenItems]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@ReceivedItemDT", GivenItems)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            GetItemsDetailsFromGivenItems = ds.Tables(0)
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return GetItemsDetailsFromGivenItems
    End Function

    Public Sub UpdateItemDescToOtherCompanies(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _ItemCode As String, ByVal _ItemDesc As String, _
                                   ByVal _Alias As String, ByVal _InActive As Boolean, ByVal _LastUpdatedBy As String, ByRef ErrNo As Integer, ByRef _ErrString As String)

        _ErrString = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("sp_UpdateItemDescToOtherCompanies", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDesc", _ItemDesc)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", _Alias)
            BaseConn.cmd.Parameters.AddWithValue("@InActive", _InActive)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", _LastUpdatedBy)
            'BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", objFromArg.dtp_LastUpdatedDate)

            'BaseConn.cmd.Parameters.AddWithValue("@DTCompanies", _DTOtherCompanies) 'BOM Parameter

            BaseConn.cmd.ExecuteNonQuery()
            'ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            '_ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            'ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function DEVTools_ItemsComparitionBetweenCompanies(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, _
                                                              ByVal _Flag As String, ByVal GivenItems As DataTable, ByRef ErrNo As Integer) As DataTable
        Dim _ErrString As String = ""
        ErrNo = 0
        DEVTools_ItemsComparitionBetweenCompanies = New DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_DEVTools_ItemsComparitionBetweenCompanies]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@GivenDT", GivenItems)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            DEVTools_ItemsComparitionBetweenCompanies = ds.Tables(0)
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return DEVTools_ItemsComparitionBetweenCompanies
    End Function

    Public Sub GetInventorySalesLedger(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _StrSiteID As String, ByVal _ItemCode As String, _
                          ByRef _InventorySalesLedger As Integer, ByRef _InventorySalesLedgerDesc As String)

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetInventorySalesLedger]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _StrSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.cmd.Parameters.Add("@InventorySalesLedger", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@InventorySalesLedgerDesc", SqlDbType.NVarChar, 100).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _InventorySalesLedger = BaseConn.cmd.Parameters("@InventorySalesLedger").Value
            _InventorySalesLedgerDesc = BaseConn.cmd.Parameters("@InventorySalesLedgerDesc").Value
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function GetInventorySalesLedgerByDONo(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _StrSiteID As String, ByVal _DTDoNo As DataTable) As DataTable
        GetInventorySalesLedgerByDONo = New DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetInventorySalesLedgerByDONo]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _StrSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@DTDoNo", _DTDoNo)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            GetInventorySalesLedgerByDONo = ds.Tables(0)
            Return GetInventorySalesLedgerByDONo
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Function

    Public Sub Get_AttributeOptions(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _StrSiteID As String, ByRef _DTAttributeOptions As DataTable, ByVal _DTAttribute As DataTable)
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetAttributeOptions]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _StrSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@DTAttribute", _DTAttribute)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _DTAttributeOptions = ds.Tables(0)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub Get_AttributeDetails(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _StrSiteID As String, ByRef _DTAttribute As DataTable, _
                                    ByRef _DTAttributeOptions As DataTable)
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetAttributeDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _StrSiteID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _DTAttribute = ds.Tables(0)
            _DTAttributeOptions = ds.Tables(1)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Get_AttributeItems(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _StrSiteID As String, ByVal _ItemDesc1 As String, ByVal _ItemDesc2 As String, ByVal _Separator1 As String, ByVal _Separator2 As String, ByVal _Separator3 As String, ByVal _Separator4 As String, ByVal _Separator5 As String, ByVal _Flag As String, ByRef _DTAttributeItem As DataTable, ByVal _ErrNo As Integer, ByRef _ErrStr As String) As DataTable
        Dim _dt As New DataTable
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetAttributeItems]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _StrSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDesc1", _ItemDesc1)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDesc2", _ItemDesc2)
            BaseConn.cmd.Parameters.AddWithValue("@Separator1", _Separator1)
            BaseConn.cmd.Parameters.AddWithValue("@Separator2", _Separator2)
            BaseConn.cmd.Parameters.AddWithValue("@Separator3", _Separator3)
            BaseConn.cmd.Parameters.AddWithValue("@Separator4", _Separator4)
            BaseConn.cmd.Parameters.AddWithValue("@Separator5", _Separator5)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@Attributes", _DTAttributeItem)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _dt = ds.Tables(0)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return _dt
    End Function

    Public Sub Update_Attribute(ByRef AttributeCode As String, ByVal obj As csAttibute, ByVal _StrSiteID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, _
                                      ByVal _Flag As String, ByRef ErrNo As Integer, ByRef _ErrString As String)

        _ErrString = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("AttributeUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _StrSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)

            BaseConn.cmd.Parameters.AddWithValue("@AttributeCode", AttributeCode)
            BaseConn.cmd.Parameters.AddWithValue("@AttributeName1", obj.str_AttributeName1)
            BaseConn.cmd.Parameters.AddWithValue("@AttributeName2", obj.str_AttributeName2)
            BaseConn.cmd.Parameters.AddWithValue("@AttributeOption", obj.bool_AttributeOption)
            BaseConn.cmd.Parameters.AddWithValue("@DT_Attribute", obj.dt_AttributeOption) 'ItemDiscList
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(objcsItemmastr.str_SiteID)
            'ObjDalGeneral.Elog_Insert(_StrSiteID, _StrDBPath, _StrDBPwd, objFromArg.int_BusinessPeriodID, objFromArg.str_CreatedBy, objFromArg.dtp_CreatedDate, "", "ItemMaster", Err.Number, "Error in " & objFromArg.str_Flag & " : " & objFromArg.str_ProductCode & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
    End Sub
End Class
