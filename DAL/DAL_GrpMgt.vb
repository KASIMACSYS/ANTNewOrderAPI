'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes
Public Class DAL_GrpMgt
    Dim objcsGrpMgt As New csGrpMgt
    Dim BaseConn As New SQLConn()
    Dim dt As New DataTable
    Private ObjDalGeneral As DAL_General

    'Public Sub Get_Structure(ByRef obj As csGrpMgt, ByVal _strPath As String, ByVal _strPwd As String, ByRef iRC As Integer, ByRef ErrStr As String)
    '    iRC = 0
    '    ErrStr = ""
    '    Try
    '        dt = New DataTable
    '        BaseConn.Open(_strPath, _strPwd)
    '        BaseConn.cmd = New SqlClient.SqlCommand("[GetGrpMgtLoad]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.Add("@CID", SqlDbType.VarChar).Value = obj.str_SiteID
    '        BaseConn.cmd.Parameters.Add("@Restriction", SqlDbType.VarChar).Value = obj.str_Restriction
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        Dim ds As New DataSet
    '        BaseConn.da.Fill(ds)

    '        obj.objcsGrpMgtCommon.dt_GrpMgtfilemain = ds.Tables(0)
    '        obj.objcsGrpMgtCommon.dt_GrpMgtfilesub = ds.Tables(1)

    '        obj.objcsGrpMgtCommon.dt_GrpMgtMastermain = ds.Tables(2)
    '        obj.objcsGrpMgtCommon.dt_GrpMgtMastersub = ds.Tables(3)

    '        obj.objcsGrpMgtCommon.dt_GrpMgtSalesMain = ds.Tables(4)
    '        obj.objcsGrpMgtCommon.dt_GrpMgtSalesSub = ds.Tables(5)

    '        obj.objcsGrpMgtCommon.dt_GrpMgtPurchaseMain = ds.Tables(6)
    '        obj.objcsGrpMgtCommon.dt_GrpMgtPurchaseSub = ds.Tables(7)

    '        obj.objcsGrpMgtCommon.dt_GrpMgtManufacturingMain = ds.Tables(8)
    '        obj.objcsGrpMgtCommon.dt_GrpMgtManufacturingSub = ds.Tables(9)

    '        obj.objcsGrpMgtCommon.dt_GrpMgtAccountsMain = ds.Tables(10)
    '        obj.objcsGrpMgtCommon.dt_GrpMgtAccountsSub = ds.Tables(11)

    '        obj.objcsGrpMgtCommon.dt_GrpMgtHRPayrollMain = ds.Tables(12)
    '        obj.objcsGrpMgtCommon.dt_GrpMgtHRPayrollSub = ds.Tables(13)

    '        obj.objcsGrpMgtCommon.dt_GrpMgtReportMain = ds.Tables(14)
    '        obj.objcsGrpMgtCommon.dt_GrpMgtReportSub = ds.Tables(15)

    '        obj.objcsGrpMgtCommon.dt_GrpMgtInventoryMain = ds.Tables(16)
    '        obj.objcsGrpMgtCommon.dt_GrpMgtInventorySub = ds.Tables(17)
    '        obj.objcsGrpMgtGS.dt_GroupMgtgeneralSettings = ds.Tables(18)

    '        obj.objcsGrpMgtCommon.dt_GrpMgtAssetMain = ds.Tables(19)
    '        obj.objcsGrpMgtCommon.dt_GrpMgtAssetSub = ds.Tables(20)

    '    Catch ex As Exception
    '        iRC = 0
    '        ErrStr = ex.Message
    '    Finally
    '        BaseConn.Close()
    '    End Try
    'End Sub
    'Public Function Get_StructureSiteID(ByVal obj As csGrpMgt, ByVal _strPath As String, ByVal _strPwd As String) As csGrpMgt
    '    Try
    '        BaseConn.Open(_strPath, _strPwd)
    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetAllSite]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        Dim ds As New DataSet
    '        BaseConn.da.Fill(ds)
    '        obj.objcsGrpMgtCommon.dt_AllSiteID = ds.Tables(0)
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    Finally
    '        BaseConn.Close()
    '    End Try
    '    Get_StructureSiteID = obj
    '    Return Get_StructureSiteID
    'End Function
    Public Function Get_StructureEdit(ByVal obj As csGrpMgt, ByVal _strPath As String, ByVal _strPwd As String) As csGrpMgt
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetGroupMgtDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.Add("@CID", SqlDbType.VarChar).Value = obj.str_SiteID
            BaseConn.cmd.Parameters.Add("@Flag", SqlDbType.VarChar).Value = obj.objcsGrpMgtCommon.str_Flag
            BaseConn.cmd.Parameters.Add("@GroupName", SqlDbType.VarChar).Value = obj.objcsGrpMgtMain.GroupName
            BaseConn.cmd.Parameters.Add("@GroupID", SqlDbType.VarChar).Value = obj.objcsGrpMgtMain.GroupID
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            obj.objcsGrpMgtMain.GroupID = ds.Tables(0).Rows(0)("GroupID").ToString
            obj.objcsGrpMgtMain.GroupSiteID = ds.Tables(0).Rows(0)("GroupSiteID").ToString
            obj.objcsGrpMgtMain.CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString
            obj.objcsGrpMgtMain.ModifiedBy = ds.Tables(0).Rows(0)("ModifiedBy").ToString
            obj.objcsGrpMgtMain.CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString
            obj.objcsGrpMgtMain.ModifiedDate = ds.Tables(0).Rows(0)("ModifiedDate").ToString
            obj.objcsGrpMgtMain.GroupLevel = ds.Tables(0).Rows(0)("GroupLevel").ToString
            'obj.objcsGrpMgtCommon.dt_GrpMgtfilemain = ds.Tables(0)
            obj.objcsGrpMgtCommon.dt_GrpMgtall = ds.Tables(1)
            obj.objcsGrpMgtCommon.dt_GrpMgtall_Report = ds.Tables(2)
            obj.objcsGrpMgtCommon.dt_groupaccpermission = ds.Tables(3)
            obj.objcsGrpMgtCommon.dt_GrpMgtSalesMan = ds.Tables(4)
            obj.objcsGrpMgtGS.dt_GroupMgtgeneralSettings = ds.Tables(5)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Get_StructureEdit = obj
        Return Get_StructureEdit
    End Function
    Public Function Update_GroupMgt(ByVal obj As csGrpMgt, ByRef GroupID As String, ByVal _strPath As String, ByVal _strPwd As String, ByRef ErrNo As String) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GroupMgtUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objcsGrpMgtCommon.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@GroupID", obj.objcsGrpMgtMain.GroupID)
            BaseConn.cmd.Parameters.AddWithValue("@GroupName", obj.objcsGrpMgtMain.GroupName)
            BaseConn.cmd.Parameters.AddWithValue("@GroupSiteID", obj.objcsGrpMgtMain.GroupSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.objcsGrpMgtMain.CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.objcsGrpMgtMain.CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ModifiedDate", obj.objcsGrpMgtMain.ModifiedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ModifiedBy", obj.objcsGrpMgtMain.ModifiedBy)
            BaseConn.cmd.Parameters.AddWithValue("@Level", obj.objcsGrpMgtMain.GroupLevel)
            BaseConn.cmd.Parameters.AddWithValue("@GroupMgtSubdt", obj.objcsGrpMgtCommon.dt_GrpMgtfilesub)
            BaseConn.cmd.Parameters.AddWithValue("@GroupMgtSubdtReport", obj.objcsGrpMgtCommon.dt_GrpMgtfilesub_Report)
            BaseConn.cmd.Parameters.AddWithValue("@GroupAccPermissiondt", obj.objcsGrpMgtCommon.dt_groupaccpermission)
            BaseConn.cmd.Parameters.AddWithValue("@GroupMgtSalesMan", obj.objcsGrpMgtCommon.dt_GrpMgtSalesMan)

            BaseConn.cmd.Parameters.AddWithValue("@GroupMgtGS", obj.objcsGrpMgtGS.dt_GroupMgtgeneralSettings)

            BaseConn.cmd.Parameters.Add("@GroupIDOut", SqlDbType.VarChar, 30).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            GroupID = BaseConn.cmd.Parameters("@GroupIDOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _strPath, _strPwd, 0, obj.objcsGrpMgtMain.CreatedBy, obj.objcsGrpMgtMain.ModifiedDate, "", "GroupMgt", Err.Number, "Error in " & obj.objcsGrpMgtCommon.str_Flag & " : " & obj.objcsGrpMgtMain.GroupName & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
        Return _ErrString
    End Function

    Public Sub Get_Structure_Report(ByRef obj As csGrpMgt, ByVal _strPath As String, ByVal _strPwd As String, ByRef iRC As Integer, ByRef ErrStr As String)
        iRC = 0
        ErrStr = ""
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetGrpMgtLoad_Report]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.Add("@SiteID", SqlDbType.VarChar).Value = obj.str_SiteID
            'BaseConn.cmd.Parameters.Add("@Flag", SqlDbType.VarChar).Value = obj.objcsGrpMgtCommon.str_Flag
            'BaseConn.cmd.Parameters.Add("@FileFrom", SqlDbType.VarChar).Value = obj.objcsGrpMgtCommon.str_FileFrom
            'BaseConn.cmd.Parameters.Add("@FileTo", SqlDbType.VarChar).Value = obj.objcsGrpMgtCommon.str_FileTo
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            obj.objcsGrpMgtCommon.dt_GrpMgtfilemain_Report = ds.Tables(0)
            obj.objcsGrpMgtCommon.dt_GrpMgtfilesub_Report = ds.Tables(1)

            obj.objcsGrpMgtCommon.dt_GrpMgtMastermain_Report = ds.Tables(2)
            obj.objcsGrpMgtCommon.dt_GrpMgtMastersub_Report = ds.Tables(3)

            obj.objcsGrpMgtCommon.dt_GrpMgtSalesMain_Report = ds.Tables(4)
            obj.objcsGrpMgtCommon.dt_GrpMgtSalesSub_Report = ds.Tables(5)

            obj.objcsGrpMgtCommon.dt_GrpMgtPurchaseMain_Report = ds.Tables(6)
            obj.objcsGrpMgtCommon.dt_GrpMgtPurchaseSub_Report = ds.Tables(7)

            obj.objcsGrpMgtCommon.dt_GrpMgtManufacturingMain_Report = ds.Tables(8)
            obj.objcsGrpMgtCommon.dt_GrpMgtManufacturingSub_Report = ds.Tables(9)

            obj.objcsGrpMgtCommon.dt_GrpMgtAccountsMain_Report = ds.Tables(10)
            obj.objcsGrpMgtCommon.dt_GrpMgtAccountsSub_Report = ds.Tables(11)

            obj.objcsGrpMgtCommon.dt_GrpMgtHRPayrollMain_Report = ds.Tables(12)
            obj.objcsGrpMgtCommon.dt_GrpMgtHRPayrollSub_Report = ds.Tables(13)

            obj.objcsGrpMgtCommon.dt_GrpMgtReportMain_Report = ds.Tables(14)
            obj.objcsGrpMgtCommon.dt_GrpMgtReportSub_Report = ds.Tables(15)

            obj.objcsGrpMgtCommon.dt_GrpMgtInventoryMain_Report = ds.Tables(16)
            obj.objcsGrpMgtCommon.dt_GrpMgtInventorySub_Report = ds.Tables(17)

            obj.objcsGrpMgtCommon.dt_GrpMgtAssetMain_Report = ds.Tables(18)
            obj.objcsGrpMgtCommon.dt_GrpMgtAssetSub_Report = ds.Tables(19)

        Catch ex As Exception
            iRC = 0
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetFormReportSettings(ByVal _strPath As String, ByVal _strPwd As String, ByRef _CID As String, ByRef _DTFormReportSettings As DataTable, ByRef iRC As Integer, ByRef ErrStr As String)
        iRC = 0
        ErrStr = ""
        Try
            _DTFormReportSettings = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetFormReportSettings]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.Add("@CID", SqlDbType.VarChar).Value = _CID
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            _DTFormReportSettings = ds.Tables(0)

        Catch ex As Exception
            iRC = 0
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetGrpMgt(ByVal _strPath As String, ByVal _strPwd As String, ByRef _SiteID As String, ByRef _GroupID As Integer, ByRef _DTGrpMgtSub As DataTable, ByRef iRC As Integer, ByRef ErrStr As String)
        iRC = 0
        ErrStr = ""
        Try
            _DTGrpMgtSub = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetGrpMgt]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.Add("@SiteID", SqlDbType.VarChar).Value = _SiteID
            BaseConn.cmd.Parameters.Add("@GroupID", SqlDbType.VarChar).Value = _GroupID
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            _DTGrpMgtSub = ds.Tables(0)

        Catch ex As Exception
            iRC = 0
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
End Class
