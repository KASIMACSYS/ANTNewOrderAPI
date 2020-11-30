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

Public Class DAL_Asset
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csAsset, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetAssetDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@AssetID", Obj.str_AssetID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", 101)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.str_AssetRefID = ds.Tables(0).Rows(0)("AssetRefID").ToString()
            Obj.str_Description = ds.Tables(0).Rows(0)("Description").ToString()
            Obj.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            Obj.int_AssetGroupCategory = ds.Tables(0).Rows(0)("AssetGroupID").ToString()
            Obj.str_Category = ds.Tables(0).Rows(0)("Category").ToString()
            Obj.str_Status = ds.Tables(0).Rows(0)("Status").ToString()
            Obj.dbl_Qty = ds.Tables(0).Rows(0)("Qty").ToString()
            Obj.int_AssetLedgerID = ds.Tables(0).Rows(0)("AssetLedgerID").ToString()
            Obj.int_AccDepLedgerID = ds.Tables(0).Rows(0)("AccDepLedgerID").ToString()
            Obj.str_BarCodeNo = ds.Tables(0).Rows(0)("BarCodeNo").ToString()
            Dim str As String = ds.Tables(0).Rows(0)("Photo").ToString()
            If str.Length > 0 Then
                Obj.img_Photo = CType(ds.Tables(0).Rows(0)("Photo"), Byte())
            Else
                Obj.img_Photo = Nothing
            End If

            Obj.str_Manufacturer = ds.Tables(0).Rows(0)("Manufacturer").ToString()
            Obj.str_Model = ds.Tables(0).Rows(0)("Model").ToString()
            Obj.str_PartNo = ds.Tables(0).Rows(0)("PartNo").ToString()
            Obj.str_SerialNo = ds.Tables(0).Rows(0)("SerialNo").ToString()
            Obj.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
            Obj.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
            Obj.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
            Obj.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
            Obj.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()

            'Obj.dtp_Manufacturing = ds.Tables(0).Rows(0)("ManufacturingDate").ToString()
            'Obj.dtp_Installation = ds.Tables(0).Rows(0)("Installation").ToString()
            'Obj.dtp_ExpiryDate = ds.Tables(0).Rows(0)("Expiry").ToString()

            'Obj.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
            'Obj.str_ContactName = ds.Tables(0).Rows(0)("ContactName").ToString()
            'Obj.str_Name = ds.Tables(0).Rows(0)("Name").ToString()
            'Obj.str_Address = ds.Tables(0).Rows(0)("Address").ToString()
            'Obj.str_Mobile = ds.Tables(0).Rows(0)("Mobile").ToString()
            'Obj.str_Tel = ds.Tables(0).Rows(0)("Tel").ToString()
            'Obj.str_Fax = ds.Tables(0).Rows(0)("Fax").ToString()
            'Obj.str_Email = ds.Tables(0).Rows(0)("Email").ToString()

            'Obj.dtp_CapitalizedOn = ds.Tables(0).Rows(0)("CapitalizedOn").ToString()
            Obj.dtp_AquisitionOn = ds.Tables(0).Rows(0)("AquisitionOn").ToString()
            Obj.dbl_AmountPosted = ds.Tables(0).Rows(0)("AmountPosted").ToString()
            Obj.str_AcqInvRef = ds.Tables(0).Rows(0)("AcqInvRef").ToString()
            Obj.int_AcqLedgerID = ds.Tables(0).Rows(0)("AcqLedgerID").ToString()
            Obj.str_AcqComment = ds.Tables(0).Rows(0)("AcqComment").ToString()

            Obj.str_DepriciationType = ds.Tables(0).Rows(0)("DepriciationType").ToString()
            Obj.dtp_StartDate = ds.Tables(0).Rows(0)("StartDate").ToString()
            Obj.dbl_DepriciationPercentage = ds.Tables(0).Rows(0)("DepriciationPercentage").ToString()
            Obj.dbl_SalvageorScrapValues = ds.Tables(0).Rows(0)("SalvageorScrapValues").ToString()
            Obj.int_NoofYears = ds.Tables(0).Rows(0)("NoofYears").ToString()
            Obj.str_DepComment = ds.Tables(0).Rows(0)("DepComment").ToString()
            Obj.int_DepLedgerID = ds.Tables(0).Rows(0)("DepLedger").ToString()

            Obj.str_DisposalType = ds.Tables(0).Rows(0)("DisposalType").ToString()
            Obj.str_DisposalInvRef = ds.Tables(0).Rows(0)("DisposalInvRef").ToString()
            Obj.dtp_DisposalDate = ds.Tables(0).Rows(0)("DisposalDate").ToString()
            Obj.int_DisposalLedgerID = ds.Tables(0).Rows(0)("DisposalLedgerID").ToString()
            Obj.int_SalesLedgerID = ds.Tables(0).Rows(0)("SalesLedgerID").ToString()
            Obj.dbl_SellingAmount = ds.Tables(0).Rows(0)("SellingAmount").ToString()
            Obj.str_DisposalComment = ds.Tables(0).Rows(0)("DisposalComment").ToString()

            Obj.XMLData = ds.Tables(0).Rows(0)("XmlData").ToString()

            Obj.dt_Transaction = ds.Tables(1)

            Obj.dt_FileUpload = ds.Tables(2)

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_Asset(ByVal obj As csAsset, ByRef _AssetNo As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef intRevNo As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("AssetUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.str_FormPrefix)

            BaseConn.cmd.Parameters.AddWithValue("@AssetID", obj.str_AssetID)
            BaseConn.cmd.Parameters.AddWithValue("@AssetRefID", obj.str_AssetRefID)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@Description", obj.str_Description)
            BaseConn.cmd.Parameters.AddWithValue("@GroupCategory", obj.int_AssetGroupCategory)
            BaseConn.cmd.Parameters.AddWithValue("@Category", obj.str_Category)
            BaseConn.cmd.Parameters.AddWithValue("@Status", obj.str_Status)
            BaseConn.cmd.Parameters.AddWithValue("@Qty", obj.dbl_Qty)
            BaseConn.cmd.Parameters.AddWithValue("@BarCodeNo", obj.str_BarCodeNo)
            If obj.img_Photo Is Nothing Then
                Dim photoParam As New SqlParameter("@Photo", SqlDbType.Image)
                photoParam.Value = DBNull.Value
                BaseConn.cmd.Parameters.Add(photoParam)
            Else
                BaseConn.cmd.Parameters.AddWithValue("@Photo", obj.img_Photo) 'PHOTO
            End If
            BaseConn.cmd.Parameters.AddWithValue("@AssetLedgerID", obj.int_AssetLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@AccDepLedgerID", obj.int_AccDepLedgerID)

            BaseConn.cmd.Parameters.AddWithValue("@Manufacturer", obj.str_Manufacturer)
            BaseConn.cmd.Parameters.AddWithValue("@Model", obj.str_Model)
            BaseConn.cmd.Parameters.AddWithValue("@PartNo", obj.str_PartNo)
            BaseConn.cmd.Parameters.AddWithValue("@SerialNo", obj.str_SerialNo)
            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.str_Comment)

            'BaseConn.cmd.Parameters.AddWithValue("@ManufacturingDate", obj.dtp_Manufacturing)
            'BaseConn.cmd.Parameters.AddWithValue("@Installation", obj.dtp_Installation)
            'BaseConn.cmd.Parameters.AddWithValue("@Expiry", obj.dtp_ExpiryDate)

            'BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.int_LedgerID)
            'BaseConn.cmd.Parameters.AddWithValue("@ContactName", obj.str_ContactName)
            'BaseConn.cmd.Parameters.AddWithValue("@Name", obj.str_Name)
            'BaseConn.cmd.Parameters.AddWithValue("@Address", obj.str_Address)
            'BaseConn.cmd.Parameters.AddWithValue("@Mobile", obj.str_Mobile)
            'BaseConn.cmd.Parameters.AddWithValue("@Tel", obj.str_Tel)
            'BaseConn.cmd.Parameters.AddWithValue("@Fax", obj.str_Fax)
            'BaseConn.cmd.Parameters.AddWithValue("@Email", obj.str_Email)

            'BaseConn.cmd.Parameters.AddWithValue("@CapitalizedOn", obj.dtp_CapitalizedOn)
            BaseConn.cmd.Parameters.AddWithValue("@AquisitionOn", obj.dtp_AquisitionOn)
            BaseConn.cmd.Parameters.AddWithValue("@AmountPosted", obj.dbl_AmountPosted)
            BaseConn.cmd.Parameters.AddWithValue("@AcqInvRef", obj.str_AcqInvRef)
            'BaseConn.cmd.Parameters.AddWithValue("@SalvageorScrapValues", obj.dbl_SalvageorScrapValues)
            BaseConn.cmd.Parameters.AddWithValue("@AcqLedgerID", obj.int_AcqLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@AcqComment", obj.str_AcqComment)


            BaseConn.cmd.Parameters.AddWithValue("@DepriciationType", obj.str_DepriciationType)
            BaseConn.cmd.Parameters.AddWithValue("@DepriciationPercentage", obj.dbl_DepriciationPercentage)
            BaseConn.cmd.Parameters.AddWithValue("@NoofYears", obj.int_NoofYears)
            BaseConn.cmd.Parameters.AddWithValue("@SalvageorScrapValues", obj.dbl_SalvageorScrapValues)
            BaseConn.cmd.Parameters.AddWithValue("@StartDate", obj.dtp_StartDate)
            BaseConn.cmd.Parameters.AddWithValue("@DepLedgerID", obj.int_DepLedgerID)
            'BaseConn.cmd.Parameters.AddWithValue("@EndDate", obj.dtp_EndDate)
            BaseConn.cmd.Parameters.AddWithValue("@DepComment", obj.str_DepComment)

            BaseConn.cmd.Parameters.AddWithValue("@DisposalType", obj.str_DisposalType)
            BaseConn.cmd.Parameters.AddWithValue("@DisposalInvRef", obj.str_DisposalInvRef)
            BaseConn.cmd.Parameters.AddWithValue("@DisposalDate", obj.dtp_DisposalDate)
            BaseConn.cmd.Parameters.AddWithValue("@DisposalLedgerID", obj.int_DisposalLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesLedgerID", obj.int_SalesLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@SellingAmount", obj.dbl_SellingAmount)
            BaseConn.cmd.Parameters.AddWithValue("@DisposalComment", obj.str_DisposalComment)

            BaseConn.cmd.Parameters.AddWithValue("@XmlData", obj.XMLData)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)

            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            BaseConn.cmd.Parameters.Add("@AssetIDOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _AssetNo = BaseConn.cmd.Parameters("@AssetIDOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, "", DateTime.Now, "", "AssetMgt", Err.Number, "Error no " & obj.str_Flag & " : " & obj.str_AssetID & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_Asset = _ErrString
    End Function

    Public Function Run_Asset(ByVal _SiteID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, _BPID As Integer, _ID As Integer, _LedgerID As Integer, _Amount As Double, _Date As Date, _Comment As String, _AssetID As String, _MenuID As String, _UserName As String, _Type As String, _Flag As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("UpdateAssetTransaction", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@AssetID", _AssetID)
            BaseConn.cmd.Parameters.AddWithValue("@BPID", _BPID)
            BaseConn.cmd.Parameters.AddWithValue("@ID", _ID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Amount", _Amount)
            BaseConn.cmd.Parameters.AddWithValue("@Date", _Date)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", _Comment)

            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@Type", _Type)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", _UserName)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Run_Asset = _ErrString
    End Function

    Public Function Update_AssetMaster(ByVal obj As csAsset, ByVal _SiteID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, _BPID As Integer, _AssetID As String, _MenuID As String, _UserName As String, _Type As String, _Flag As String, _ID As String, ByRef ErrNo As Integer) As String
        Dim _ErrStr As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("UpdateAssetMaster", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@AssetID", _AssetID)
            BaseConn.cmd.Parameters.AddWithValue("@BPID", _BPID)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", _UserName)
            BaseConn.cmd.Parameters.AddWithValue("@Type", _Type)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@ID", _ID)

            BaseConn.cmd.Parameters.AddWithValue("@AquisitionOn", obj.dtp_AquisitionOn)
            BaseConn.cmd.Parameters.AddWithValue("@AmountPosted", obj.dbl_AmountPosted)
            BaseConn.cmd.Parameters.AddWithValue("@AcqInvRef", obj.str_AcqInvRef)
            BaseConn.cmd.Parameters.AddWithValue("@AcqLedgerID", obj.int_AcqLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@AcqComment", obj.str_AcqComment)

            BaseConn.cmd.Parameters.AddWithValue("@DisposalType", obj.str_DisposalType)
            BaseConn.cmd.Parameters.AddWithValue("@DisposalInvRef", obj.str_DisposalInvRef)
            BaseConn.cmd.Parameters.AddWithValue("@DisposalDate", obj.dtp_DisposalDate)
            BaseConn.cmd.Parameters.AddWithValue("@DisposalLedgerID", obj.int_DisposalLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesLedgerID", obj.int_SalesLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@SellingAmount", obj.dbl_SellingAmount)
            BaseConn.cmd.Parameters.AddWithValue("@DisposalComment", obj.str_DisposalComment)

            BaseConn.cmd.Parameters.AddWithValue("@ReValDate", obj.dtp_ReValDate)
            BaseConn.cmd.Parameters.AddWithValue("@ReValLedgerID", obj.int_ReValLedger)
            BaseConn.cmd.Parameters.AddWithValue("@ReValAmount", obj.dbl_ReValAmount)
            BaseConn.cmd.Parameters.AddWithValue("@ReValComment", obj.str_ReValComment)

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrStr = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrStr = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_AssetMaster = _ErrStr
    End Function

End Class
