'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes
Public Class DAL_EmployeeSettings
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General


    Public Sub Get_Structure(ByRef ObjEmpSettings As csEmployeeSettings, ByRef dt_MccbSalesMan As DataTable, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal ErrNo As String, ByVal ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetSalesmanSettingsDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", ObjEmpSettings.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", ObjEmpSettings.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", ObjEmpSettings.str_SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", ObjEmpSettings.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            If ObjEmpSettings.str_Flag = "ALL" Then
                dt_MccbSalesMan = ds.Tables(0)
            Else
                ObjEmpSettings.str_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()
                ObjEmpSettings.str_SalesManName = ds.Tables(0).Rows(0)("SalesManName").ToString()
                ObjEmpSettings.str_Alias1 = ds.Tables(0).Rows(0)("Alias1").ToString()
                ObjEmpSettings.str_Alias2 = ds.Tables(0).Rows(0)("Alias2").ToString()
                ObjEmpSettings.str_EmployeeLedgerID = ds.Tables(0).Rows(0)("EmployeeLedgerID").ToString()
                ObjEmpSettings.dbl_PaymentLimit = ds.Tables(0).Rows(0)("PaymentLimit").ToString()
                ObjEmpSettings.int_LimitStatus = ds.Tables(0).Rows(0)("LimitStatus").ToString()
                ObjEmpSettings.dbl_Commission = ds.Tables(0).Rows(0)("Commission").ToString()
                ObjEmpSettings.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                ObjEmpSettings.bool_InActive = ds.Tables(0).Rows(0)("InActive").ToString()
            End If
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub Update_EmployeeSettings(ByVal ObjEmployeeSettings As csEmployeeSettings, ByRef _SalesManID As String, ByRef SiteID As String, ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrStr = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[SalesmanSettingsUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", ObjEmployeeSettings.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", ObjEmployeeSettings.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", ObjEmployeeSettings.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@EmployeeLedgerID", ObjEmployeeSettings.str_EmployeeLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", ObjEmployeeSettings.str_SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManName", ObjEmployeeSettings.str_SalesManName)
            BaseConn.cmd.Parameters.AddWithValue("@Alias1", ObjEmployeeSettings.str_Alias1)
            BaseConn.cmd.Parameters.AddWithValue("@Alias2", ObjEmployeeSettings.str_Alias2)
            BaseConn.cmd.Parameters.AddWithValue("@PaymentLimit", ObjEmployeeSettings.dbl_PaymentLimit)
            BaseConn.cmd.Parameters.AddWithValue("@LimitStatus", ObjEmployeeSettings.int_LimitStatus)
            BaseConn.cmd.Parameters.AddWithValue("@Commission", ObjEmployeeSettings.dbl_Commission)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", ObjEmployeeSettings.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@InActive", ObjEmployeeSettings.bool_InActive)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", ObjEmployeeSettings.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", ObjEmployeeSettings.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", ObjEmployeeSettings.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", ObjEmployeeSettings.dtp_LastUpdatedDate)

            BaseConn.cmd.Parameters.Add("@SalesManIDOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()

            _SalesManID = BaseConn.cmd.Parameters("@SalesManIDOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            ErrStr = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
            ObjDalGeneral = New DAL_General(ObjEmployeeSettings.int_CID)
            ObjDalGeneral.Elog_Insert(ObjEmployeeSettings.str_SalesManID, _strDBPath, _StrDBPwd, 0, ObjEmployeeSettings.str_CreatedBy, ObjEmployeeSettings.dtp_LastUpdatedDate, ObjEmployeeSettings.str_CreatedBy, "EmpVacation", Err.Number, "Error in " & ObjEmployeeSettings.str_Flag & " : " & ObjEmployeeSettings.str_SalesManName & "", ex.Message, 5, 3, 1, ErrNo)
        Finally
            BaseConn.Close()
        End Try
    End Sub

End Class