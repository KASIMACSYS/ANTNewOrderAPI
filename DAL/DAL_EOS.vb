Imports Classes

Public Class DAL_EOS
    Dim BaseConn As New SQLConn()

    Public Sub GetEOS(ByVal strSiteID As String, ByVal strDBPath As String, ByVal strDBPWD As String, ByVal strEOSNO As String, ByRef dtMain As DataTable, ByRef dtSub As DataTable, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            dtMain = New DataTable
            dtSub = New DataTable
            BaseConn.Open(strDBPath, strDBPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetEOS]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", strSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@EOSID", strEOSNO)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dtMain = ds.Tables(0)
            dtSub = ds.Tables(1)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetLeaveDatails(ByVal strSiteID As String, ByVal strDBPath As String, ByVal strDBPWD As String, ByVal intLedgerID As Integer, ByVal strCategory As String, ByVal dtpDate As Date, ByVal strFlag As String, ByRef dt As DataTable, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            dt = New DataTable
            BaseConn.Open(strDBPath, strDBPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_EOSCalculation]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", strSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", intLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Category", strCategory)
            BaseConn.cmd.Parameters.AddWithValue("@EndDate", dtpDate)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", strFlag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            If strFlag = "LEAVESALARY" Then
                dt = ds.Tables(1)
            Else
                dt = ds.Tables(3)
            End If

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub UpdateEOS(ByVal strSiteID As String, ByVal strDBPath As String, ByVal strDBPWD As String, ByVal strEOSNO As String, ByVal intLedgerID As Integer, ByVal dtpEndDate As Date, ByVal dtpSalaryMonth As Date, ByVal strFlag As String, ByVal dt As DataTable, ByRef ErrNo As Integer, ByRef ErrStr As String, ByRef OutEOSNO As String, ByRef OutRevNo As Integer, ByVal strPrefix As String, ByVal intRevNo As Integer, ByVal MenuID As String, ByVal UserName As String)
        ErrNo = 0
        ErrStr = ""
        Try
            'dt = New DataTable
            BaseConn.Open(strDBPath, strDBPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_EOSUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", strSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@EOSID", strEOSNO)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", intLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@EndMonth", dtpEndDate)
            BaseConn.cmd.Parameters.AddWithValue("@SalaryMonth", dtpSalaryMonth)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", strFlag)
            BaseConn.cmd.Parameters.AddWithValue("@Prefix", strPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", intRevNo)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", MenuID)

            BaseConn.cmd.Parameters.AddWithValue("@EOSDT", dt)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", UserName)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", UserName)


            BaseConn.cmd.Parameters.Add("@EOSIDOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output

            BaseConn.cmd.ExecuteNonQuery()

            OutEOSNO = BaseConn.cmd.Parameters("@EOSIDOut").Value.ToString
            OutRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString

            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            ErrStr = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub

End Class
