Imports Classes
Public Class DAL_SalesManGrouping
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Public Sub Get_Structure(ByRef Obj As csSalesManGrouping, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[SalesManGroupingUpdated]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", Obj.str_SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@Description", Obj.str_Description)
            BaseConn.cmd.Parameters.AddWithValue("@ParentID", Obj.str_ParentID)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", Obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", Obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", Obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", Obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManIDDT", Obj.dt_SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub Get_Structure1(ByRef Obj As csSalesManGrouping, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetSalesManGrouping]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", Obj.str_SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Obj.Condition)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.dt_selecteditem = ds.Tables(0)
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub
End Class
