'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes
Public Class DAL_CurrencyMaster
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csCurrencyMaster, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrString As String)
        ErrNo = 0
        ErrString = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            dt = New DataTable
            BaseConn.cmd = New SqlClient.SqlCommand("[GetCurrencyMasterDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CurrencyID", Obj.str_CurrencyID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            If Obj.str_Flag = "GET" Then
                Obj.str_CurrencyCode = dt.Rows(0)("CurrencyCode").ToString()
                Obj.str_CurrencyName = dt.Rows(0)("CurrencyName").ToString()
                Obj.bool_DefaultCurrency = dt.Rows(0)("BaseCurrencyFlag").ToString()
                Obj.dbl_PurExchangeRate = dt.Rows(0)("PurExchangeRate").ToString()
                Obj.dbl_ExchangeRate = dt.Rows(0)("ExchangeRate").ToString()
                Obj.dbl_SalExchangeRate = dt.Rows(0)("SalExchangeRate").ToString()
                Obj.int_DecimalPlace = dt.Rows(0)("DecimalPlace").ToString()
                Obj.str_MajorCurrency = dt.Rows(0)("MajorCurrencyText").ToString()
                Obj.str_MinorCurrency = dt.Rows(0)("MinorCurrencytext").ToString()

                Obj.str_CreatedBy = dt.Rows(0)("CreatedBy").ToString()
                Obj.dtp_CreatedDate = dt.Rows(0)("CreatedDate").ToString()
                Obj.str_LastUpdatedBy = dt.Rows(0)("LastUpdatedBy").ToString()
                Obj.dtp_LastUpdatedDate = dt.Rows(0)("LastUpdatedDate").ToString()
                Obj.str_ApprovedBy = dt.Rows(0)("ApprovedBy").ToString()
                Obj.dtp_ApprovedDate = dt.Rows(0)("ApprovedDate").ToString()
                Obj.bool_ApprovedStatus = dt.Rows(0)("ApprovedStatus")
            End If
        Catch ex As Exception
            ErrNo = 1
            ErrString = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Function Update_CurrencyMaster(ByVal obj As csCurrencyMaster, ByRef str_VouNo As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("CurrencyMasterUpdated", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CurrencyID", obj.str_CurrencyID)
            BaseConn.cmd.Parameters.AddWithValue("@CurrencyCode", obj.str_CurrencyCode)
            BaseConn.cmd.Parameters.AddWithValue("@CurrencyName", obj.str_CurrencyName)
            BaseConn.cmd.Parameters.AddWithValue("@BaseCurrencyFlag", obj.bool_DefaultCurrency)
            BaseConn.cmd.Parameters.AddWithValue("@PurExchangeRate", obj.dbl_PurExchangeRate)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.dbl_ExchangeRate)
            BaseConn.cmd.Parameters.AddWithValue("@SalExchangeRate", obj.dbl_SalExchangeRate)

            BaseConn.cmd.Parameters.AddWithValue("@DecimalPlace", obj.int_DecimalPlace)
            BaseConn.cmd.Parameters.AddWithValue("@MajorCurrencyText", obj.str_MajorCurrency)
            BaseConn.cmd.Parameters.AddWithValue("@MinorCurrencyText", obj.str_MinorCurrency)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "CurrencyMaster", Err.Number, "Error in " & obj.str_Flag & " : " & obj.str_CurrencyID & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
        Update_CurrencyMaster = _ErrString
    End Function
End Class
