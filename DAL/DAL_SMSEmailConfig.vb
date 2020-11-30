Imports Classes

Public Class DAL_SMSEmailConfig
    Private BaseConn As New SQLConn()
    Private dt As DataTable
    Private ObjDalGeneral As DAL_General

    Public Function Get_Structure(ByRef Obj As csSMSEmailConfig, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As Integer,
                                 ByVal _Flag As String, ByRef _ErrNo As Integer, ByRef _ErrString As String) As DataTable
        dt = New DataTable
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetSMSEmailConfigDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
            Obj.dt_SMSEmail = dt
            'Obj.str_LanguageCode = ds.Tables(0).Rows(0)("RevNo").ToString()
            'Obj.str_TextMsg = ds.Tables(0).Rows(0)("Comment").ToString()

        Catch ex As Exception
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub Put_Structure(ByRef Obj As csSMSEmailConfig, ByRef _VouNo As String, ByRef _RevNo As Integer, ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        _ErrString = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("SMSEmailConfigUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", Obj.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", Obj.str_LanguageCode)
            BaseConn.cmd.Parameters.AddWithValue("@SMSorEmail", Obj.str_SMSorEmail)
            BaseConn.cmd.Parameters.AddWithValue("@TextMsg", Obj.str_TextMsg)
            BaseConn.cmd.Parameters.AddWithValue("@AddorEditFlag", Obj.str_AddorEditFlag)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", Obj.CreatedBy)

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output

            BaseConn.cmd.ExecuteNonQuery()
            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(Obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(Obj.str_SiteID, _DBPath, _DBPwd, Obj.int_BusinessPeriodID, Obj.str_CreatedBy, Obj.dtp_CreatedDate, "", "SMSEmailConfig", Err.Number, "Error in '" & Obj.str_Flag & "'ED '" & Obj.str_SMSorEmail & "' ", ex.Message, 5, 3, 1, 0)
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

End Class
