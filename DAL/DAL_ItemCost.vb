
Imports Classes
Public Class DAL_ItemCost
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csItemCost, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetItemCostDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@DocumentNo", Obj.ObjItemCostMain.str_DocumentNo)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.ObjItemCostMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.ObjItemCostMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            Obj.ObjItemCostMain.dtp_DocumentDate = ds.Tables(0).Rows(0)("Date").ToString()
            Obj.ObjItemCostMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()


            If ds.Tables(1).Rows.Count > 0 Then
                Obj.ObjItemCostSub.dt_CostTypeSub = ds.Tables(1)
            End If

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub


    Public Function Update_ItemCost(ByVal obj As csItemCost, ByRef str_DocumentNo As String, ByRef intRevNo As Integer, ByVal _StrDBPath As String, ByVal _strDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("ItemCostUpdated", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.ObjItemCostMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@DocumentNo", obj.ObjItemCostMain.str_DocumentNo)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.ObjItemCostMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@Date", obj.ObjItemCostMain.dtp_DocumentDate)
            'BaseConn.cmd.Parameters.AddWithValue("@DocDate", obj.ObjItemCostMain.dtp_DocumentDate)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.ObjItemCostMain.str_Comment)

            'BaseConn.cmd.Parameters.AddWithValue("@WHID", obj.ObjItemCostMain.str_WHID)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.ObjItemCostMain.str_ItemCostPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@ItemCostDT", obj.ObjItemCostSub.dt_CostTypeSub)
            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", obj.ObjItemCostMain.int_LanguageCode)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjItemCostMain.str_Flag)
            BaseConn.cmd.Parameters.Add("@DocumentNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.cmd.ExecuteNonQuery()
            str_DocumentNo = BaseConn.cmd.Parameters("@DocumentNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _strDBPwd, 0, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "CostAdjustment", Err.Number, "Error in " & obj.ObjItemCostMain.str_Flag & " : " & obj.ObjItemCostMain.str_DocumentNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_ItemCost = _ErrString
    End Function

End Class
