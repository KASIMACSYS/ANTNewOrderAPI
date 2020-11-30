
Imports Classes
Public Class DAL_Bin
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Function Update_Bin(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal obj As csBinMaster, ByVal _Flag As String, _
                                 ByVal _BusinessPeriodID As Integer, ByRef intRevNo As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0


        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("UpdateBinMaster", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)

            BaseConn.cmd.Parameters.AddWithValue("@BinID", obj.str_BinID)
            BaseConn.cmd.Parameters.AddWithValue("@BinDesc", obj.str_BinDesc)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@InActive", obj.bool_InActive)
            BaseConn.cmd.Parameters.AddWithValue("@WH_ID", obj.str_WHID)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString

        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, _BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "Batch", _
             Err.Number, "Error in " & _Flag & " : " & obj.str_BinID & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_Bin = _ErrString
    End Function

    Public Sub Get_Structure(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef Obj As csBinMaster, ByRef ErrNo As Integer, _
                             ByRef ErrStr As String)

        ErrNo = 0
        ErrStr = ""

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetBinMasterDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BinID", Obj.str_BinID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.str_BinDesc = ds.Tables(0).Rows(0)("BinDesc").ToString()
            Obj.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            Obj.str_WHID = ds.Tables(0).Rows(0)("WH_ID").ToString()
            Obj.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.bool_InActive = ds.Tables(0).Rows(0)("InActive").ToString()

            Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
            Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
            Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
            Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
            Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_BINMovement(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal str_SiteID As String, ByVal obj As csBinMovement, ByRef str_DocumentNo As String, ByRef outRevNo As Integer, ByRef ErrNo As Integer, ByRef ErrStr As String) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("BinMovementUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.ObjBinMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@RefID", obj.ObjBinMain.str_RefID)
            BaseConn.cmd.Parameters.AddWithValue("@RefDate", obj.ObjBinMain.dtp_RefDate)
            BaseConn.cmd.Parameters.AddWithValue("@FromWH", obj.ObjBinMain.str_FromWH)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.ObjBinMain.str_Comment)

            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.ObjBinMain.int_RevNo)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.ObjBinMain.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.ObjBinMain.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.ObjBinMain.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.ObjBinMain.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.ObjBinMain.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.ObjBinMain.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.ObjBinMain.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@DT", obj.ObjBinSub.dt_BinSub)

            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjBinMain.str_Flag)
            BaseConn.cmd.Parameters.Add("@RefIDOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            str_DocumentNo = BaseConn.cmd.Parameters("@RefIDOut").Value.ToString
            outRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.ObjBinMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "BinMovement", Err.Number, "Error in '" & obj.ObjBinMain.str_Flag & "'ED '" & obj.ObjBinMain.str_RefID & "' ", ex.Message, 5, 3, 1, 0)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_BINMovement = _ErrString
    End Function
    Public Sub Get_BinMovement(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal str_SiteID As String, ByVal Obj As csBinMovement, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetBinMovementDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@RefID", Obj.ObjBinMain.str_RefID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.ObjBinMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.ObjBinMain.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.ObjBinMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString
            Obj.ObjBinMain.dtp_RefDate = ds.Tables(0).Rows(0)("RefDate").ToString
            Obj.ObjBinMain.str_FromWH = ds.Tables(0).Rows(0)("WHID").ToString()
            Obj.ObjBinMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.ObjBinMain.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
            Obj.ObjBinMain.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
            Obj.ObjBinMain.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.ObjBinMain.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            Obj.ObjBinMain.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
            Obj.ObjBinMain.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
            Obj.ObjBinMain.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus").ToString()
            If ds.Tables(1).Rows.Count > 0 Then
                Obj.ObjBinSub.dt_BinSub = ds.Tables(1)
            End If
            'Obj.DTBatch = ds.Tables(2)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
        'Get_BinMovement = Obj
        'Return Get_BinMovement()
    End Sub
End Class