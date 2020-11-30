Imports Classes

Public Class DAL_Logistics
    Private ObjDalGeneral As DAL_General
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()



    Public Sub Get_Structure(ByRef Obj As csLogistics, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetLogisticsDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@VoucherNo", Obj.str_VoucherNo)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.int_RevNo = ds.Tables(0).Rows(0)("RevNo")
            Obj.dtp_VoucherDate = ds.Tables(0).Rows(0)("VoucherDate")
            Obj.str_DoNo = ds.Tables(0).Rows(0)("DoNo").ToString()
            Obj.str_CusName = ds.Tables(0).Rows(0)("CusName").ToString()
            Obj.dtp_DoDate = ds.Tables(0).Rows(0)("DoDate")
            Obj.str_DeliveryLocation = ds.Tables(0).Rows(0)("DeliveryLocation").ToString()
            Obj.str_DOCreadtedBy = ds.Tables(0).Rows(0)("DOCreatedBy").ToString()
            Obj.bool_DOStatus = ds.Tables(0).Rows(0)("DoStatus").ToString()
            Obj.str_TruckNo = ds.Tables(0).Rows(0)("TruckNo").ToString()
            Obj.str_DeliverName = ds.Tables(0).Rows(0)("DeliverName").ToString()
            Obj.str_MobileNo = ds.Tables(0).Rows(0)("MobileNo").ToString()
            Obj.dtp_TimePrint = ds.Tables(0).Rows(0)("TimePrint").ToString()

            Obj.dbl_CargoCharges = ds.Tables(0).Rows(0)("CargoCharges").ToString()
            Obj.bool_TransportStatus = ds.Tables(0).Rows(0)("TransportStatus").ToString()
            Obj.str_GatePass = ds.Tables(0).Rows(0)("GatePass").ToString()
            Obj.str_CustomRef = ds.Tables(0).Rows(0)("CustomRef").ToString()
            Obj.dbl_CustomCharges = ds.Tables(0).Rows(0)("CustomCharges").ToString()
            Obj.str_AirwayBillNo = ds.Tables(0).Rows(0)("AirwayBillNo").ToString()
            Obj.dtp_ExitDate = ds.Tables(0).Rows(0)("ExitDate")
            Obj.bool_CustomStatus = ds.Tables(0).Rows(0)("CustomStatus").ToString()
            Obj.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_Logistics(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal obj As csLogistics, ByRef VouNo As String, ByRef intRevNo As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("sp_LogisticsUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@VoucherNo", obj.str_VoucherNo)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@VoucherDate", obj.dtp_VoucherDate)
            BaseConn.cmd.Parameters.AddWithValue("@DoNo", obj.str_DoNo)
            BaseConn.cmd.Parameters.AddWithValue("@CusName", obj.str_CusName)
            BaseConn.cmd.Parameters.AddWithValue("@DoDate", obj.dtp_DoDate)
            BaseConn.cmd.Parameters.AddWithValue("@DeliveryLocation", obj.str_DeliveryLocation)
            BaseConn.cmd.Parameters.AddWithValue("@DoCreatedBy", obj.str_DOCreadtedBy)
            BaseConn.cmd.Parameters.AddWithValue("@DoStatus", obj.bool_DOStatus)
            BaseConn.cmd.Parameters.AddWithValue("@TruckNo", obj.str_TruckNo)
            BaseConn.cmd.Parameters.AddWithValue("@DeliverName", obj.str_DeliverName)
            BaseConn.cmd.Parameters.AddWithValue("@MobileNo", obj.str_MobileNo)
            BaseConn.cmd.Parameters.AddWithValue("@TimePrint", obj.dtp_TimePrint)
            BaseConn.cmd.Parameters.AddWithValue("@CargoCharges", obj.dbl_CargoCharges)
            BaseConn.cmd.Parameters.AddWithValue("@TransportStatus", obj.bool_TransportStatus)
            BaseConn.cmd.Parameters.AddWithValue("@GatePass", obj.str_GatePass)
            BaseConn.cmd.Parameters.AddWithValue("@CustomRef", obj.str_CustomRef)
            BaseConn.cmd.Parameters.AddWithValue("@CustomCharges", obj.dbl_CustomCharges)
            BaseConn.cmd.Parameters.AddWithValue("@AirwayBillNo", obj.str_AirwayBillNo)
            BaseConn.cmd.Parameters.AddWithValue("@ExitDate", obj.dtp_ExitDate)
            BaseConn.cmd.Parameters.AddWithValue("@CustomStatus", obj.bool_CustomStatus)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            VouNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPerionID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "Logistics", Err.Number, "Error in " & obj.str_Flag & " : " & obj.str_VoucherNo & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_Logistics = _ErrString
    End Function
End Class
