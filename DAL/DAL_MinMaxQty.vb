
Imports Classes
Public Class DAL_MinMaxQty
    Private BaseConn As New SQLConn()
    Private dt As DataTable
    Private ObjDalGeneral As DAL_General
    Public Function GetItemWiseLedger(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal int_BusinessPeroidID As Integer, ByVal Str_MinMaxQty As String, ByVal dtp_Date As Date, ByVal condition As String, Optional ByVal ItemCodeColl As DataTable = Nothing, Optional ByVal VendorColl As DataTable = Nothing, Optional ByVal _WHID As Integer = 0) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)

            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetMinMaxQty]", BaseConn.cnn)

            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPeroidID)
            BaseConn.cmd.Parameters.AddWithValue("@MinMaxQty", Str_MinMaxQty)
            BaseConn.cmd.Parameters.AddWithValue("@DtpDate", dtp_Date)
            BaseConn.cmd.Parameters.AddWithValue("@ItemArray", ItemCodeColl)
            BaseConn.cmd.Parameters.AddWithValue("@VendorArray", VendorColl)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", condition)
            If Not _WHID = 0 Then
                BaseConn.cmd.Parameters.AddWithValue("@WHID", _WHID)
            End If
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
End Class
