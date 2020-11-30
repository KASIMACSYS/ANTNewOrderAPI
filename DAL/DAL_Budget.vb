Imports Classes

Public Class DAL_Budget
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef obj As csBudget, ByVal _StrDBPath As String, ByVal str_SiteID As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String, ByVal _Flag As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetBudgetDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@BudgetID", obj.ObjBudgetMain.Str_BudgetID)
            BaseConn.cmd.Parameters.AddWithValue("@GroupLedger", obj.ObjBudgetMain.str_GroupLedger)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            If _Flag = "ADD" Then
                obj.ObjBudgetMain.dt_Budget = ds.Tables(0)
            Else
                obj.ObjBudgetMain.int_RevNo = ds.Tables(1).Rows(0)("RevNo").ToString()
                obj.ObjBudgetMain.Str_BudgetID = ds.Tables(1).Rows(0)("BudgetID").ToString()
                obj.ObjBudgetMain.str_Description = ds.Tables(1).Rows(0)("Description").ToString()
                obj.ObjBudgetMain.dtp_VouDate = ds.Tables(1).Rows(0)("VoucherDate").ToString()
                obj.ObjBudgetMain.dtp_FromDate = ds.Tables(1).Rows(0)("FromMonth").ToString()
                obj.ObjBudgetMain.dtp_ToDate = ds.Tables(1).Rows(0)("ToMonth").ToString()
                obj.ObjBudgetMain.Str_Comment = ds.Tables(1).Rows(0)("Comment").ToString()
                obj.ObjBudgetMain.Str_Status = ds.Tables(1).Rows(0)("Status").ToString()
                obj.ObjBudgetMain.dt_Budget = ds.Tables(0)
            End If
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_Budget(ByVal obj As csBudget, ByRef VouNo As String, ByRef intRevNo As Integer, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("BudgetUpdated", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Prefix", obj.ObjBudgetMain.Str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjBudgetMain.Str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.ObjBudgetMain.Str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@BudgetID", obj.ObjBudgetMain.Str_BudgetID)
            BaseConn.cmd.Parameters.AddWithValue("@Description", obj.ObjBudgetMain.str_Description)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.ObjBudgetMain.Str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@VoucherDate", obj.ObjBudgetMain.dtp_VouDate)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", obj.ObjBudgetMain.dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", obj.ObjBudgetMain.dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Status", obj.ObjBudgetMain.Str_Status)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)

            BaseConn.cmd.Parameters.AddWithValue("@BudgetDT", obj.ObjBudgetMain.dt_Budget)

            BaseConn.cmd.Parameters.Add("@BudgetIDOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500

            BaseConn.cmd.ExecuteNonQuery()

            VouNo = BaseConn.cmd.Parameters("@BudgetIDOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "Budget", Err.Number, "Error in " & obj.ObjBudgetMain.Str_Flag & " : " & obj.ObjBudgetMain.Str_BudgetID & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_Budget = _ErrString
    End Function

    Public Function Get_Report(ByRef obj As csBudget, ByVal BudgetID As String, ByVal _StrDBPath As String, ByVal str_SiteID As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String, ByVal _Flag As String)
        ErrNo = 0
        ErrStr = ""
        Try
            dt = New DataTable
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[BudgetReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@BudgetID", BudgetID)
            BaseConn.cmd.Parameters.AddWithValue("@Status", obj.ObjBudgetMain.Str_Status)
            BaseConn.cmd.Parameters.AddWithValue("@Options", BudgetID)
            BaseConn.cmd.Parameters.AddWithValue("@DateType", obj.ObjBudgetMain.str_DateType)
            BaseConn.cmd.Parameters.AddWithValue("@Date", obj.ObjBudgetMain.dtp_Date)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", obj.ObjBudgetMain.dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", obj.ObjBudgetMain.dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@SignatureType", "")
            BaseConn.cmd.Parameters.AddWithValue("@User", obj.ObjBudgetMain.Str_User)
            BaseConn.cmd.Parameters.AddWithValue("@AccountingPeriod", obj.ObjBudgetMain.int_AccountingPeriodFrom)
            BaseConn.cmd.Parameters.AddWithValue("@AccountingPeriodFrom", obj.ObjBudgetMain.str_AccountingPeriod)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
End Class
