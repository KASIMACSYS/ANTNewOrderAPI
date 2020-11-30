Imports Classes
Public Class DAL_MAInvoiceDetails
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()

    Public Function Get_MAInvoiceDetails(ByVal str_SiteID As String, ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal str_MenuID As String, ByVal str_Flag As String, ByVal dtp_FromDate As DateTime, ByVal dtp_ToDate As DateTime, ByRef ErrNo As Integer, ByRef ErrStr As String) As DataTable
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_MAGetInvoiceDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_ToDate)
            dt = New DataTable
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            ErrStr = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
End Class
