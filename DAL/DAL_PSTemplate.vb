Imports Classes
Public Class DAL_PSTemplate
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Function Put_Structure(ByRef Obj As csPSTemplate, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef int_TemplateID As String, ByRef ErrNo As Integer, ByRef ErrStr As String) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("PSTemplateUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", Obj.str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", Obj.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Prefix", Obj.str_Prefix)
            BaseConn.cmd.Parameters.AddWithValue("@Description", Obj.str_Description)
            BaseConn.cmd.Parameters.AddWithValue("@Interval", Obj.str_Interval)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", Obj.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", Obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", Obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", Obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", Obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", Obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", Obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", Obj.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@DT", Obj.dt_PSTemplatesub)
            BaseConn.cmd.Parameters.AddWithValue("@VouNoOut", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            int_TemplateID = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Put_Structure = _ErrString
    End Function

    Public Sub Get_Structure(ByRef Obj As csPSTemplate, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetPSTemplateDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@Description", Obj.str_Description)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", Obj.str_VouNo)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.str_VouNo = ds.Tables(0).Rows(0)("VouNo").ToString()
            Obj.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            Obj.str_Description = ds.Tables(0).Rows(0)("Description").ToString()
            Obj.str_Interval = ds.Tables(0).Rows(0)("Interval").ToString()
            Obj.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.dt_PSTemplatesub = ds.Tables(1)
        Catch ex As Exception
            ErrStr = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

End Class
