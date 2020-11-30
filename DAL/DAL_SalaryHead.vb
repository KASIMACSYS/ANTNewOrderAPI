Imports Classes
Public Class DAL_SalaryHead
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csSalaryHead, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetSalaryHead]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPerionID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", "")
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.dt_Salaryhead = ds.Tables(0)
            Obj.dt_Source = ds.Tables(1)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
   
    Public Function Update_SalaryHead(ByVal obj As csSalaryHead, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_SalaryHead]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPerionID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", "")
            BaseConn.cmd.Parameters.AddWithValue("@SalaryHeadDT", obj.dt_Salaryhead)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPerionID, obj.str_CreatedBy, DateTime.Now, "", "SalaryHead", Err.Number, "Error in : " & obj.str_Flag & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_SalaryHead = _ErrString
    End Function
End Class
