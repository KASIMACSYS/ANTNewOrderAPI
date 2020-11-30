'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes

Public Class DAL_EmpVacation

    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef ObjEmpVacDetails As csEmpVacation, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal ErrNo As String, ByVal ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetEmpVacationDetails]", BaseConn.cnn) 'sp_GetEmpVacation
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", ObjEmpVacDetails.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", ObjEmpVacDetails.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", ObjEmpVacDetails.str_Flag) '= "DGV_EmpVacDetails"
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", ObjEmpVacDetails.int_LedgerID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            ObjEmpVacDetails.dt_EmpVacation.Clear()
            BaseConn.da.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                ObjEmpVacDetails.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                ObjEmpVacDetails.str_VacDetails = ds.Tables(0).Rows(0)("VacDetails").ToString()
                ObjEmpVacDetails.dtp_From = ds.Tables(0).Rows(0)("FromDate").ToString()
                ObjEmpVacDetails.dtp_To = ds.Tables(0).Rows(0)("ToDate").ToString()
                ObjEmpVacDetails.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                ObjEmpVacDetails.str_RtnComment = ds.Tables(0).Rows(0)("RtnComment").ToString()
                ObjEmpVacDetails.dtp_Today = ds.Tables(0).Rows(0)("PostedDate").ToString()
                ObjEmpVacDetails.dtp_RtnDate = ds.Tables(0).Rows(0)("RtnDate").ToString

                ObjEmpVacDetails.dt_EmpVacation = ds.Tables(0)
            End If
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_EmpVacDetails(ByVal ObjEmpVacDetails As csEmpVacation, ByRef SiteID As String, ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[EmpVacationUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", ObjEmpVacDetails.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", ObjEmpVacDetails.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", ObjEmpVacDetails.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", ObjEmpVacDetails.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@PostedDate", ObjEmpVacDetails.dtp_Today)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", ObjEmpVacDetails.dtp_From)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate ", ObjEmpVacDetails.dtp_To)
            BaseConn.cmd.Parameters.AddWithValue("@RtnDate", ObjEmpVacDetails.dtp_RtnDate)
            BaseConn.cmd.Parameters.AddWithValue("@VacDetails", ObjEmpVacDetails.str_VacDetails)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", ObjEmpVacDetails.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@RtnComment", ObjEmpVacDetails.str_RtnComment)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", ObjEmpVacDetails.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", ObjEmpVacDetails.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", ObjEmpVacDetails.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", ObjEmpVacDetails.dtp_LastUpdatedDate)

            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", ObjEmpVacDetails.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", ObjEmpVacDetails.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", ObjEmpVacDetails.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()


            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(ObjEmpVacDetails.int_CID)
            'ObjDalGeneral.Elog_Insert(obj.str_SiteID, _strDBPath, _StrDBPwd, 0, obj.str_UserName, obj.dtp_LastUpdatedDate, "", "UserMgt", Err.Number, "Error in " & obj.str_Flag & " : " & obj.str_UserName & "  ", ex.Message, 5, 3, 1, ErrNo)
            'ObjDalGeneral.Elog_Insert(ObjEmpVacDetails.str_SiteID, _StrDBPwd, 0, ObjEmpVacDetails.int_LedgerID, ObjEmpVacDetails.dtp_LastUpdatedDate, "", "EmpVacation", Err.Number, "Error in " & ObjEmpVacDetails.str_Flag & " : ", ex.Message, 5, 3, 1, ErrNo)
            ObjDalGeneral.Elog_Insert(ObjEmpVacDetails.int_CID, _strDBPath, _StrDBPwd, 0, ObjEmpVacDetails.str_CreatedBy, ObjEmpVacDetails.dtp_LastUpdatedDate, ObjEmpVacDetails.str_CreatedBy, "EmpVacation", Err.Number, "Error in " & ObjEmpVacDetails.str_Flag & " : " & ObjEmpVacDetails.int_LedgerID & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_EmpVacDetails = _ErrString
    End Function
End Class
