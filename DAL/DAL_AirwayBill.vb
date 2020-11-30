
Imports Classes

Public Class DAL_AirwayBill
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByRef Obj As csAirwayBill, ByRef ErrNo As Integer, ByRef ErrMsg As String)
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetAirwayBillDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objAirwayBillMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", Obj.objAirwayBillMain.str_RefNo)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            'Obj.objAirwayBillMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
            Obj.objAirwayBillMain.dtp_Date = ds.Tables(0).Rows(0)("VouDate")
            Obj.objAirwayBillMain.str_Text1 = ds.Tables(0).Rows(0)("Text1").ToString()
            Obj.objAirwayBillMain.str_Text2 = ds.Tables(0).Rows(0)("Text2").ToString()
            Obj.objAirwayBillMain.str_Text3 = ds.Tables(0).Rows(0)("Text3").ToString()
            Obj.objAirwayBillMain.str_ShipperName = ds.Tables(0).Rows(0)("ShipperName").ToString()
            Obj.objAirwayBillMain.str_IssuedBy = ds.Tables(0).Rows(0)("IssuedBy").ToString()
            Obj.objAirwayBillMain.str_ConsignName = ds.Tables(0).Rows(0)("ConsignName").ToString()
            Obj.objAirwayBillMain.str_IssueAgent = ds.Tables(0).Rows(0)("IssueAgent").ToString()
            Obj.objAirwayBillMain.str_AccInfo = ds.Tables(0).Rows(0)("AccInfo").ToString()
            Obj.objAirwayBillMain.str_AirportDept = ds.Tables(0).Rows(0)("AirportDept").ToString()
            Obj.objAirwayBillMain.str_AirportTo = ds.Tables(0).Rows(0)("AirportTo").ToString()
            Obj.objAirwayBillMain.str_ByFirstCarrier = ds.Tables(0).Rows(0)("ByFirstCarrier").ToString()
            Obj.objAirwayBillMain.str_Currency = ds.Tables(0).Rows(0)("Currency").ToString()
            Obj.objAirwayBillMain.str_ChqCode = ds.Tables(0).Rows(0)("ChqCode").ToString()
            Obj.objAirwayBillMain.str_WTVALPPD = ds.Tables(0).Rows(0)("WTVALPPD").ToString()
            Obj.objAirwayBillMain.str_WTVALCOLL = ds.Tables(0).Rows(0)("WTVALCOLL").ToString()
            Obj.objAirwayBillMain.str_OtherPPD = ds.Tables(0).Rows(0)("OtherPPD").ToString()
            Obj.objAirwayBillMain.str_OtherCOLL = ds.Tables(0).Rows(0)("OtherCOLL").ToString()
            Obj.objAirwayBillMain.str_ValueForCarriage = ds.Tables(0).Rows(0)("ValueForCarriage").ToString()
            Obj.objAirwayBillMain.str_ValueForCustoms = ds.Tables(0).Rows(0)("ValueForCustoms").ToString()
            Obj.objAirwayBillMain.str_AirportofDest = ds.Tables(0).Rows(0)("AirportofDest").ToString()
            Obj.objAirwayBillMain.str_HandlingInfo = ds.Tables(0).Rows(0)("HandlingInfo").ToString()
            Obj.objAirwayBillMain.str_TotalPCS = ds.Tables(0).Rows(0)("TotalPCS").ToString()
            Obj.objAirwayBillMain.str_TotalGrossWeight = ds.Tables(0).Rows(0)("TotalGrossWeight").ToString()
            Obj.objAirwayBillMain.str_Total = ds.Tables(0).Rows(0)("Total").ToString()
            Obj.objAirwayBillMain.str_WCPrepaid = ds.Tables(0).Rows(0)("WCPrepaid").ToString()
            Obj.objAirwayBillMain.str_WCCollect = ds.Tables(0).Rows(0)("WCCollect").ToString()
            Obj.objAirwayBillMain.str_ValueChargePrepaid = ds.Tables(0).Rows(0)("ValueChargePrepaid").ToString()
            Obj.objAirwayBillMain.str_ValueChargeCollect = ds.Tables(0).Rows(0)("ValueChargeCollect").ToString()

            Obj.objAirwayBillMain.str_TaxPrepaid = ds.Tables(0).Rows(0)("TaxPrepaid").ToString()
            Obj.objAirwayBillMain.str_TaxCollect = ds.Tables(0).Rows(0)("TaxCollect").ToString()
            Obj.objAirwayBillMain.str_DueAgentPrepaid = ds.Tables(0).Rows(0)("DueAgentPrepaid").ToString()
            Obj.objAirwayBillMain.str_DueAgentCollect = ds.Tables(0).Rows(0)("DueAgentCollect").ToString()
            Obj.objAirwayBillMain.str_DueCarrierPrepaid = ds.Tables(0).Rows(0)("DueCarrierPrepaid").ToString()
            Obj.objAirwayBillMain.str_DueCarrierCollect = ds.Tables(0).Rows(0)("DueCarrierCollect").ToString()
            Obj.objAirwayBillMain.str_OtherCharge = ds.Tables(0).Rows(0)("OtherCharge").ToString()
            Obj.objAirwayBillMain.str_TotalPrepaid = ds.Tables(0).Rows(0)("TotalPrepaid").ToString()
            Obj.objAirwayBillMain.str_TotalCollect = ds.Tables(0).Rows(0)("TotalCollect").ToString()
            Obj.objAirwayBillMain.str_ExecDate = ds.Tables(0).Rows(0)("ExecDate").ToString()
            Obj.objAirwayBillMain.str_ExecPlace = ds.Tables(0).Rows(0)("ExecPlace").ToString()
            Obj.objAirwayBillMain.str_Extra1 = ds.Tables(0).Rows(0)("Extra1").ToString()

            Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
            Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
            Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
            Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
            Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")

            Obj.DT_AirwayBillSub = ds.Tables(1)
            'If ds.Tables(1).Rows.Count > 0 Then
            '    Obj.DT_AirwayBillSub = ds.Tables(1)
            'End If

        Catch ex As Exception
            ErrNo = 1
            ErrMsg = ex.Message ' "Problem in Updating Invoice"
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Put_Structure(ByVal _strPath As String, ByVal _strPwd As String, ByVal obj As csAirwayBill, ByRef RefNo As String, ByRef intRevNo As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("sp_AirwayBillUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objAirwayBillMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objAirwayBillMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objAirwayBillMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.objAirwayBillMain.str_FormPrefix)

            BaseConn.cmd.Parameters.AddWithValue("@VouDate", obj.objAirwayBillMain.dtp_Date)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", obj.objAirwayBillMain.str_RefNo)
            BaseConn.cmd.Parameters.AddWithValue("@Text1", obj.objAirwayBillMain.str_Text1)
            BaseConn.cmd.Parameters.AddWithValue("@Text2", obj.objAirwayBillMain.str_Text2)
            BaseConn.cmd.Parameters.AddWithValue("@Text3", obj.objAirwayBillMain.str_Text3)
            BaseConn.cmd.Parameters.AddWithValue("@ShipperName", obj.objAirwayBillMain.str_ShipperName)

            BaseConn.cmd.Parameters.AddWithValue("@IssuedBy", obj.objAirwayBillMain.str_IssuedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ConsignName", obj.objAirwayBillMain.str_ConsignName)
            BaseConn.cmd.Parameters.AddWithValue("@IssueAgent", obj.objAirwayBillMain.str_IssueAgent)
            BaseConn.cmd.Parameters.AddWithValue("@AccInfo", obj.objAirwayBillMain.str_AccInfo)
            BaseConn.cmd.Parameters.AddWithValue("@AirportDept", obj.objAirwayBillMain.str_AirportDept)
            BaseConn.cmd.Parameters.AddWithValue("@AirportTo", obj.objAirwayBillMain.str_AirportTo)
            BaseConn.cmd.Parameters.AddWithValue("@ByFirstCarrier", obj.objAirwayBillMain.str_ByFirstCarrier)
            BaseConn.cmd.Parameters.AddWithValue("@Currency", obj.objAirwayBillMain.str_Currency)
            BaseConn.cmd.Parameters.AddWithValue("@ChqCode", obj.objAirwayBillMain.str_ChqCode)
            BaseConn.cmd.Parameters.AddWithValue("@WTVALPPD", obj.objAirwayBillMain.str_WTVALPPD)
            BaseConn.cmd.Parameters.AddWithValue("@WTVALCOLL", obj.objAirwayBillMain.str_WTVALCOLL)
            BaseConn.cmd.Parameters.AddWithValue("@OtherPPD", obj.objAirwayBillMain.str_OtherPPD)
            BaseConn.cmd.Parameters.AddWithValue("@OtherCOLL", obj.objAirwayBillMain.str_OtherCOLL)
            BaseConn.cmd.Parameters.AddWithValue("@ValueForCarriage", obj.objAirwayBillMain.str_ValueForCarriage)
            BaseConn.cmd.Parameters.AddWithValue("@ValueForCustoms", obj.objAirwayBillMain.str_ValueForCustoms)
            BaseConn.cmd.Parameters.AddWithValue("@AirportofDest", obj.objAirwayBillMain.str_AirportofDest)
            BaseConn.cmd.Parameters.AddWithValue("@HandlingInfo", obj.objAirwayBillMain.str_HandlingInfo)
            BaseConn.cmd.Parameters.AddWithValue("@TotalPcs", obj.objAirwayBillMain.str_TotalPCS)

            BaseConn.cmd.Parameters.AddWithValue("@TotalGrossWeight", obj.objAirwayBillMain.str_TotalGrossWeight)
            BaseConn.cmd.Parameters.AddWithValue("@Total", obj.objAirwayBillMain.str_Total)
            BaseConn.cmd.Parameters.AddWithValue("@WCPrepaid", obj.objAirwayBillMain.str_WCPrepaid)
            BaseConn.cmd.Parameters.AddWithValue("@WCCollect", obj.objAirwayBillMain.str_WCCollect)
            BaseConn.cmd.Parameters.AddWithValue("@ValueChargePrepaid", obj.objAirwayBillMain.str_ValueChargePrepaid)
            BaseConn.cmd.Parameters.AddWithValue("@ValueChargeCollect", obj.objAirwayBillMain.str_ValueChargeCollect)
            BaseConn.cmd.Parameters.AddWithValue("@TaxPrepaid", obj.objAirwayBillMain.str_TaxPrepaid)
            BaseConn.cmd.Parameters.AddWithValue("@TaxCollect", obj.objAirwayBillMain.str_TaxCollect)
            BaseConn.cmd.Parameters.AddWithValue("@DueAgentPrepaid", obj.objAirwayBillMain.str_DueAgentPrepaid)
            BaseConn.cmd.Parameters.AddWithValue("@DueAgentCollect", obj.objAirwayBillMain.str_DueAgentCollect)
            BaseConn.cmd.Parameters.AddWithValue("@DueCarrierPrepaid", obj.objAirwayBillMain.str_DueCarrierPrepaid)
            BaseConn.cmd.Parameters.AddWithValue("@DueCarrierCollect", obj.objAirwayBillMain.str_DueCarrierCollect)
            BaseConn.cmd.Parameters.AddWithValue("@OtherCharge", obj.objAirwayBillMain.str_OtherCharge)
            BaseConn.cmd.Parameters.AddWithValue("@TotalPrepaid", obj.objAirwayBillMain.str_TotalPrepaid)
            BaseConn.cmd.Parameters.AddWithValue("@TotalCollect", obj.objAirwayBillMain.str_TotalCollect)

            BaseConn.cmd.Parameters.AddWithValue("@ExecDate", obj.objAirwayBillMain.str_ExecDate)
            BaseConn.cmd.Parameters.AddWithValue("@ExecPlace", obj.objAirwayBillMain.str_ExecPlace)
            BaseConn.cmd.Parameters.AddWithValue("@Extra1", obj.objAirwayBillMain.str_Extra1)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@AirwayBillDT", obj.DT_AirwayBillSub)

            BaseConn.cmd.Parameters.Add("@AirWayBillOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            'BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            RefNo = BaseConn.cmd.Parameters("@AirWayBillOut").Value.ToString
            'intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _strPath, _strPwd, obj.objAirwayBillMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "AirwayBill", ErrNo, "Error in " & obj.objAirwayBillMain.str_Flag & " : " & obj.objAirwayBillMain.str_RefNo & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Put_Structure = _ErrString
    End Function

  
End Class
