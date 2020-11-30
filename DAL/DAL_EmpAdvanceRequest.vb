Imports Classes

Public Class DAL_EmpAdvanceRequest
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General
    Public Sub Get_Structure(ByRef obj As csAdvanceRequest, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetEmpAdvanceRequestDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjAdvanceRequestMain.Str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@ReqID", obj.ObjAdvanceRequestMain.Str_ReqID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            obj.ObjAdvanceRequestMain.Str_ReqID = ds.Tables(0).Rows(0)("ReqID").ToString()
            obj.ObjAdvanceRequestMain.int_EmpLedgerID = ds.Tables(0).Rows(0)("EmpLedgerID").ToString()
            obj.ObjAdvanceRequestMain.Str_Type = ds.Tables(0).Rows(0)("Type").ToString()
            obj.ObjAdvanceRequestMain.Str_Description = ds.Tables(0).Rows(0)("Description").ToString()
            obj.ObjAdvanceRequestMain.dbl_AmountRequested = ds.Tables(0).Rows(0)("AmountRequested").ToString()
            obj.ObjAdvanceRequestMain.Str_Currency = ds.Tables(0).Rows(0)("Currency").ToString()
            obj.ObjAdvanceRequestMain.bool_Approved = ds.Tables(0).Rows(0)("Approved").ToString()
            obj.ObjAdvanceRequestMain.dbl_ApprovedAmount = ds.Tables(0).Rows(0)("ApprovedAmount").ToString()
            obj.ObjAdvanceRequestMain.dbl_MonthlyDeduct = ds.Tables(0).Rows(0)("MonthlyDeduct").ToString()
            obj.ObjAdvanceRequestMain.int_DeductMonthCount = ds.Tables(0).Rows(0)("DeductionMonthCount").ToString()
            obj.ObjAdvanceRequestMain.dtp_StartDate = ds.Tables(0).Rows(0)("StartDate").ToString()
            If ds.Tables(1).Rows.Count > 0 Then
                obj.ObjAdvanceRequestMain.Str_PaymentReqID = ds.Tables(1).Rows(0)("ReqID").ToString()
            Else
                obj.ObjAdvanceRequestMain.Str_PaymentReqID = ""
            End If
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub


    Public Function Update_EmpRequestLeave(ByVal obj As csAdvanceRequest, ByRef VouNo As String, ByRef intRevNo As Integer, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("EmpAdvanceRequestUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.ObjAdvanceRequestMain.Str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjAdvanceRequestMain.Str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.ObjAdvanceRequestMain.Str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@ReqID", obj.ObjAdvanceRequestMain.Str_ReqID)
            BaseConn.cmd.Parameters.AddWithValue("@EmpLedgerID", obj.ObjAdvanceRequestMain.int_EmpLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Type", obj.ObjAdvanceRequestMain.Str_Type)
            BaseConn.cmd.Parameters.AddWithValue("@Description", obj.ObjAdvanceRequestMain.Str_Description)
            BaseConn.cmd.Parameters.AddWithValue("@AmountRequest", obj.ObjAdvanceRequestMain.dbl_AmountRequested)
            BaseConn.cmd.Parameters.AddWithValue("@CurrencyCode", obj.ObjAdvanceRequestMain.Str_Currency)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedAmount", obj.ObjAdvanceRequestMain.dbl_ApprovedAmount)
            BaseConn.cmd.Parameters.AddWithValue("@Approved", obj.ObjAdvanceRequestMain.bool_Approved)
            BaseConn.cmd.Parameters.AddWithValue("@Status", obj.ObjAdvanceRequestMain.Str_Status)
            BaseConn.cmd.Parameters.AddWithValue("@MonthlyDeduct", obj.ObjAdvanceRequestMain.dbl_MonthlyDeduct)
            BaseConn.cmd.Parameters.AddWithValue("@DeductionMonthCount", obj.ObjAdvanceRequestMain.int_DeductMonthCount)
            BaseConn.cmd.Parameters.AddWithValue("@StartDate", obj.ObjAdvanceRequestMain.dtp_StartDate)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.ObjAdvanceRequestMain.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.ObjAdvanceRequestMain.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.ObjAdvanceRequestMain.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.ObjAdvanceRequestMain.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.ObjAdvanceRequestMain.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.ObjAdvanceRequestMain.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.Add("@ReqIDOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            VouNo = BaseConn.cmd.Parameters("@ReqIDOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_CID)
            ObjDalGeneral.Elog_Insert(obj.str_CID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "EmpAdvanceRequest", Err.Number, "Error in " & obj.ObjAdvanceRequestMain.Str_Flag & " : " & obj.ObjAdvanceRequestMain.Str_ReqID & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_EmpRequestLeave = _ErrString
    End Function
End Class
