
Imports Classes

Public Class DAL_MultiApprovalLedger
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()

    Public Sub Get_structure(ByVal obj As csMultiApprovalLedger, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetMultiApprovalLedger]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.str_BusinessPerionID)
            BaseConn.cmd.Parameters.AddWithValue("@Module", obj.str_Module)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.str_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@DocumentType", obj.str_DocumentType)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", obj.str_User)
            BaseConn.cmd.Parameters.AddWithValue("@GroupID", obj.str_GroupID)
            BaseConn.cmd.Parameters.AddWithValue("@AccountPeriod", obj.str_ACPeriod)
            BaseConn.cmd.Parameters.AddWithValue("@DateType", obj.str_DateType)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", obj.dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", obj.dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@OrderBy", obj.str_ApprovedByOrder)
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)
            obj.dt_Main = dt
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub

End Class
