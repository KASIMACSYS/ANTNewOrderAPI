Imports Classes

Public Class DAL_WPS
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Function GetDTForWPSExcel(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal ObjHRMain As csHRMain, ByVal Query As String, _
                                     ByRef ErrNo As Integer, ByRef ErrStr As String) As DataTable

        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            dt = New DataTable
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetWPSReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", ObjHRMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", ObjHRMain.str_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Category", ObjHRMain.str_Category)
            BaseConn.cmd.Parameters.AddWithValue("@Date", ObjHRMain._Date)
            BaseConn.cmd.Parameters.AddWithValue("@FrmDate", ObjHRMain.dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", ObjHRMain.dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Query", Query)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
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
