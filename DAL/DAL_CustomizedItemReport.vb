'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes
Public Class DAL_CustomizedItemReport
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Public Sub Get_Structure(ByRef Obj As csCustomizedItemReport, ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrString As String)
        ErrNo = 0
        ErrString = ""
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetCustomizedItemReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.str_BusinessPerionID)
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", Obj.ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@ItemArray", Obj.dt_ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", Obj.dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", Obj.dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Date1", Obj.dtp_Date)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantID", Obj.str_MerchantID)
            BaseConn.cmd.Parameters.AddWithValue("@WHLocation", Obj.WHLocation)
            BaseConn.cmd.Parameters.AddWithValue("@ZeroSuppress", Obj.bool_ZeroSuppress)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.dt_Main = ds.Tables(0)
        Catch ex As Exception
            ErrNo = 1
            ErrString = ex.Message
        Finally
            BaseConn.Close()
        End Try
        'Return Get_Structure
    End Sub
    Public Sub Get_CustomizedVoucherSearch(ByRef Obj As csCustomizedItemReport, ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrString As String)
        ErrNo = 0
        ErrString = ""
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetCustomizedVoucherSearch]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.str_BusinessPerionID)
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesMan", Obj.objCustomizedVoucher.str_SalesMan)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", Obj.dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", Obj.dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Date1", Obj.dtp_Date)
            BaseConn.cmd.Parameters.AddWithValue("@QTN", Obj.objCustomizedVoucher.bool_QTN)
            BaseConn.cmd.Parameters.AddWithValue("@SalOrd", Obj.objCustomizedVoucher.bool_SalOrd)
            BaseConn.cmd.Parameters.AddWithValue("@DO", Obj.objCustomizedVoucher.bool_DO)
            BaseConn.cmd.Parameters.AddWithValue("@SIS", Obj.objCustomizedVoucher.bool_SIS)
            BaseConn.cmd.Parameters.AddWithValue("@SRT", Obj.objCustomizedVoucher.bool_SRT)
            BaseConn.cmd.Parameters.AddWithValue("@ISSUE", Obj.objCustomizedVoucher.bool_ISSUE)
            BaseConn.cmd.Parameters.AddWithValue("@GIP", Obj.objCustomizedVoucher.bool_GIP)
            BaseConn.cmd.Parameters.AddWithValue("@LPO", Obj.objCustomizedVoucher.bool_LPO)
            BaseConn.cmd.Parameters.AddWithValue("@MRV", Obj.objCustomizedVoucher.bool_MRV)
            BaseConn.cmd.Parameters.AddWithValue("@PIP", Obj.objCustomizedVoucher.bool_PIP)
            BaseConn.cmd.Parameters.AddWithValue("@PRT", Obj.objCustomizedVoucher.bool_PRT)
            BaseConn.cmd.Parameters.AddWithValue("@GEP", Obj.objCustomizedVoucher.bool_GEP)
            BaseConn.cmd.Parameters.AddWithValue("@RVCash", Obj.objCustomizedVoucher.bool_RVCash)
            BaseConn.cmd.Parameters.AddWithValue("@RVCheq", Obj.objCustomizedVoucher.bool_RVCheq)
            BaseConn.cmd.Parameters.AddWithValue("@PVCash", Obj.objCustomizedVoucher.bool_PVCash)
            BaseConn.cmd.Parameters.AddWithValue("@PVCheq", Obj.objCustomizedVoucher.bool_PVCheq)
            BaseConn.cmd.Parameters.AddWithValue("@GEV", Obj.objCustomizedVoucher.bool_GEV)
            BaseConn.cmd.Parameters.AddWithValue("@JV", Obj.objCustomizedVoucher.bool_JV)
            BaseConn.cmd.Parameters.AddWithValue("@PRODUCTION", Obj.objCustomizedVoucher.bool_PRODUCTION)
            BaseConn.cmd.Parameters.AddWithValue("@WildSearch", Obj.objCustomizedVoucher.str_WildSerach)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.dt_Main = ds.Tables(0)
        Catch ex As Exception
            ErrNo = 1
            ErrString = ex.Message
        Finally
            BaseConn.Close()
        End Try
        'Return Get_Structure
    End Sub
End Class
