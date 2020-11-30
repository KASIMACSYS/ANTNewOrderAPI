Imports Classes

Public Class DAL_CustomsInvoice
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef Obj As csCustomsInvoice, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetCIDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CINo", Obj.objCIMain.str_CINo)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objCIMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.objCIMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
            Obj.objCIMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            Obj.objCIMain.dtp_Date = ds.Tables(0).Rows(0)("CIDate").ToString
            Obj.objCIMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
            Obj.objCIMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
            Obj.objCIMain.str_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()
            Obj.objCIMain.dbl_ItemDiscount = ds.Tables(0).Rows(0)("ItemDiscount").ToString()

            Obj.objCIMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.objCIMain.str_Cargo = ds.Tables(0).Rows(0)("Cargo").ToString()
            Obj.objCIMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
            Obj.objCIMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()

            Obj.objCIMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
            Obj.objCIMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
            Obj.objCIMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
            Obj.objCIMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
            Obj.objCIMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()


            Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
            Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
            Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
            Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
            Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")

            Obj.DT_CIItemDetails = ds.Tables(1)

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub


    Public Function Update_CI(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal obj As csCustomsInvoice, ByRef VouNo As String, _
                              ByRef intRevNo As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("CIUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.objCIMain.str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objCIMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objCIMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objCIMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.objCIMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@CINo", obj.objCIMain.str_CINo)
            BaseConn.cmd.Parameters.AddWithValue("@CIDate", obj.objCIMain.dtp_Date)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.objCIMain.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.objCIMain.str_Alias)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", obj.objCIMain.str_SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objCIMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@Cargo", obj.objCIMain.str_Cargo)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDiscount", obj.objCIMain.dbl_ItemDiscount)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.objCIMain.str_TCCurrency)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.objCIMain.dbl_ExchangeRate)

            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.objCIMain.dbl_TCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", obj.objCIMain.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", obj.objCIMain.dbl_TCDiscountAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", obj.objCIMain.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", obj.objCIMain.dbl_LCNetAmount)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@CIItemDetailsDT", obj.DT_CIItemDetails)
            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            VouNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString

        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.objCIMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, _
                                      "", "DO", Err.Number, "Error in " & obj.objCIMain.str_Flag & " : " & obj.objCIMain.str_CINo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_CI = _ErrString
    End Function
End Class
