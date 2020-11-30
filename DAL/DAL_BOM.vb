
Imports Classes

Public Class DAL_BOM
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByRef Obj As csBOM, ByRef ErrNo As Integer, ByRef ErrMsg As String)
        Try
            ErrNo = 0
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetBOMDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BOMNo", Obj.objBOMMain.str_BOMNo)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objBOMMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_CID)
            'BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.objBOMMain.str_Flag)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.objBOMMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
            Obj.objBOMMain.str_BOMNo = ds.Tables(0).Rows(0)("BOMNo").ToString()
            Obj.objBOMMain.dtp_Date = ds.Tables(0).Rows(0)("BOMDate").ToString()
            Obj.objBOMMain.str_BOMDesc = ds.Tables(0).Rows(0)("BOMDesc").ToString()
            Obj.objBOMMain.str_ItemCode = ds.Tables(0).Rows(0)("ItemCode").ToString()
            Obj.objBOMMain.str_ItemDesc = ds.Tables(0).Rows(0)("ItemDesc").ToString()
            Obj.objBOMMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            Obj.objBOMMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()


            Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
            Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
            Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
            Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
            Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")

            If ds.Tables(1).Rows.Count > 0 Then
                Obj.DT_BOMItemDetails = ds.Tables(1)
            End If

            If ds.Tables(2).Rows.Count > 0 Then
                Obj.DT_BOMParameters = ds.Tables(2)
            End If

        Catch ex As Exception
            ErrNo = 1
            ErrMsg = ex.Message ' "Problem in Updating Invoice"
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_BOM(ByVal _strPath As String, ByVal _strPwd As String, ByVal obj As csBOM, ByRef BOMNo As String, ByRef intRevNo As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("BOMUpdated", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objBOMMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objBOMMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objBOMMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.objBOMMain.str_FormPrefix)

            BaseConn.cmd.Parameters.AddWithValue("@BOMDate", obj.objBOMMain.dtp_Date)
            BaseConn.cmd.Parameters.AddWithValue("@BOMNo", obj.objBOMMain.str_BOMNo)
            BaseConn.cmd.Parameters.AddWithValue("@BOMDesc", obj.objBOMMain.str_BOMDesc)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", obj.objBOMMain.str_ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDesc", obj.objBOMMain.str_ItemDesc)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objBOMMain.str_Comment)


            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@BOMItemDetailsDT", obj.DT_BOMItemDetails)
            BaseConn.cmd.Parameters.AddWithValue("@BOMParametersDT", obj.DT_BOMParameters)

            BaseConn.cmd.Parameters.Add("@BOMNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            BOMNo = BaseConn.cmd.Parameters("@BOMNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_CID)
            ObjDalGeneral.Elog_Insert(obj.str_CID, _strPath, _strPwd, obj.objBOMMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "BOM", ErrNo, "Error in " & obj.objBOMMain.str_Flag & " : " & obj.objBOMMain.str_BOMNo & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_BOM = _ErrString
    End Function

    Public Sub Get_BOMItemDetails(ByVal _DBPath As String, ByVal _DBPwd As String, ByRef Obj As csBOM, ByVal _ItemCode As String, ByRef ErrNo As Integer, ByRef ErrMsg As String)
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetBOMItemDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BOMNo", Obj.objBOMMain.str_BOMNo)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objBOMMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            'If ds.Tables(1).Rows.Count > 0 Then
            Obj.DT_BOMItemDetails = ds.Tables(0)
            'End If

            'If ds.Tables(2).Rows.Count > 0 Then
            Obj.DT_BOMParameters = ds.Tables(1)
            'End If

        Catch ex As Exception
            ErrNo = 1
            ErrMsg = ex.Message ' "Problem in Updating Invoice"
        Finally
            BaseConn.Close()
        End Try
    End Sub
End Class
