Imports Classes
Public Class DAL_WpsReportConfig
    Dim dt As DataTable
    Dim Basecon As New SQLConn()
    Private ObjDalGeneral As DAL_General
    Public Function put_structure(ByVal obj As csWpsReportConfig, ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            Basecon.Open(_strDBPath, _StrDBPwd)
            Basecon.cmd = New SqlClient.SqlCommand("sp_WpsReportConfigUpdate", Basecon.cnn)
            Basecon.cmd.CommandType = CommandType.StoredProcedure
            Basecon.cmd.Parameters.AddWithValue("@SiteID", obj.str_SiteID) 'obj.str_SiteID
            Basecon.cmd.Parameters.AddWithValue("@WpsID", obj.str_WpsId)
            Basecon.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            Basecon.cmd.Parameters.AddWithValue("@DT", obj.dt)
            Basecon.cmd.Parameters.AddWithValue("@HRDT", obj.dt_Sub)

            Basecon.cmd.Parameters.Add("@EmpIDOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            Basecon.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            Basecon.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            Basecon.cmd.ExecuteNonQuery()
            ErrNo = Basecon.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = Basecon.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _strDBPath, _StrDBPwd, obj.int_BusinessPeriodID, "", Date.Now, "", "WpsReportConfig", Err.Number, "Error in " & obj.str_Flag & " : " & obj.str_WpsId & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            Basecon.Close()
        End Try

        put_structure = _ErrString

    End Function
    Public Function Get_structure(ByVal obj As csWpsReportConfig, ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal ErrNo As Integer, ByVal ErrStr As String) As csWpsReportConfig
        Try
            Basecon.Open(_strDBPath, _StrDBPwd)
            Basecon.cmd = New SqlClient.SqlCommand("sp_GetWpsReportConfig", Basecon.cnn)
            Basecon.cmd.CommandType = CommandType.StoredProcedure
            Basecon.cmd.Parameters.AddWithValue("@SiteID", obj.str_SiteID)
            Basecon.cmd.Parameters.AddWithValue("@WPSID", obj.str_WpsId)
            Basecon.da = New SqlClient.SqlDataAdapter(Basecon.cmd)
            Dim ds As New DataSet
            Basecon.da.Fill(ds)
            obj.dt = ds.Tables(0)
            obj.dt_Sub = ds.Tables(1)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            Basecon.Close()
        End Try
        Get_structure = obj
        Return Get_structure
    End Function

End Class
