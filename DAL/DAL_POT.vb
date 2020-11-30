Imports Classes

Public Class DAL_POT
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csPOT, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[FE_GetPOTDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", Obj.ObjPOTMain.Str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", "")
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.ObjPOTMain.dtp_VouDate = ds.Tables(0).Rows(0)("VouDate").ToString()
            Obj.ObjPOTMain.Str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.ObjPOTMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo")
            Obj.ObjPOTMain.str_JONo = ds.Tables(0).Rows(0)("JONo")
            Obj.ObjPOTMain.str_ItemCode = ds.Tables(0).Rows(0)("ItemCode")
            Obj.ObjPOTMain.str_ItemDesc = ds.Tables(0).Rows(0)("Alias1")
            Obj.ObjPOTMain.int_OrgSlno = ds.Tables(0).Rows(0)("OrgSlno")
            Obj.ObjPOTMain.str_Status = ds.Tables(0).Rows(0)("ItemStatus")
            Obj.ObjPOTMain.str_BOQ = ds.Tables(0).Rows(0)("BOQ")
            Obj.ObjPOTMain.str_SONo = ds.Tables(0).Rows(0)("SONo")
            Obj.ObjPOTMain.str_DrawingNo = ds.Tables(0).Rows(0)("DrawingNo")
            Obj.ObjPOTMain.str_Area = ds.Tables(0).Rows(0)("Area")
            Obj.ObjPOTMain.str_Unit = ds.Tables(0).Rows(0)("Unit")
            Obj.ObjPOTMain.str_Qty = ds.Tables(0).Rows(0)("VouQty")
            Obj.ObjPOTMain.str_Price = ds.Tables(0).Rows(0)("Price")
            Obj.ObjPOTMain.str_Amount = ds.Tables(0).Rows(0)("Amount")
            Obj.ObjPOTMain.str_POTNo = ds.Tables(0).Rows(0)("POTNo")
            Obj.ObjPOTMain.str_POTDesc = ds.Tables(0).Rows(0)("POTDesc").ToString()
            Obj.ObjPOTMain.dtp_StartDate = ds.Tables(0).Rows(0)("StartDate").ToString()
            Obj.ObjPOTMain.dtp_EndDate = ds.Tables(0).Rows(0)("EndDate").ToString()

            Obj.ObjPOTMain.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy")
            Obj.ObjPOTMain.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
            Obj.ObjPOTMain.dtp_IssuedDate = ds.Tables(0).Rows(0)("IssuedDate").ToString()
            Obj.ObjPOTMain.str_IssuedBy = ds.Tables(0).Rows(0)("IssuedBy").ToString()
            Obj.ObjPOTMain.str_ReceivedBy = ds.Tables(0).Rows(0)("ReceivedBy").ToString()
            Obj.ObjPOTMain.str_FactoryTeam = ds.Tables(0).Rows(0)("FactoryTeam").ToString()
            Obj.ObjPOTMain.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel")
            Obj.ObjPOTMain.int_RevisionNo = ds.Tables(0).Rows(0)("RevisionNo")

            'Obj.LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            'Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            'Obj.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
            ''Obj.ObjStkAdjMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()
            Obj.ObjPOTSub.DT_POT = ds.Tables(1)
            Obj.ObjPOTSub.dt_Attachment = ds.Tables(2)
            Obj.ObjPOTSub.dt_Section = ds.Tables(3)
            If ds.Tables(4).Rows.Count > 0 Then
                Obj.objproject.str_ProjectID = ds.Tables(4).Rows(0)("ProjectID").ToString()
                Obj.objproject.str_ProjectLocation = ds.Tables(4).Rows(0)("ProjectLocation").ToString()
                Obj.objproject.str_WorkOrderNo = ds.Tables(4).Rows(0)("WorkOrderNo").ToString()
            Else
                Obj.objproject.str_ProjectID = ""
                Obj.objproject.str_ProjectLocation = ""
                Obj.objproject.str_WorkOrderNo = ""
            End If
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_POT(ByVal obj As csPOT, ByRef str_DocumentNo As String, ByRef intRevNo As Integer, ByVal _StrDBPath As String, ByVal _strDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[FE_POTUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.ObjPOTMain.Str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", obj.ObjPOTMain.Str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@VouDate", obj.ObjPOTMain.dtp_VouDate)
            BaseConn.cmd.Parameters.AddWithValue("@StartDate", obj.ObjPOTMain.dtp_StartDate)
            BaseConn.cmd.Parameters.AddWithValue("@EndDate", obj.ObjPOTMain.dtp_EndDate)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.ObjPOTMain.Str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.ObjPOTMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@JONo", obj.ObjPOTMain.str_JONo)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", obj.ObjPOTMain.str_ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDesc", obj.ObjPOTMain.str_ItemDesc)
            BaseConn.cmd.Parameters.AddWithValue("@OrgSlno", obj.ObjPOTMain.int_OrgSlno)
            BaseConn.cmd.Parameters.AddWithValue("@Status", obj.ObjPOTMain.str_Status)
            BaseConn.cmd.Parameters.AddWithValue("@BOQ", obj.ObjPOTMain.str_BOQ)
            BaseConn.cmd.Parameters.AddWithValue("@SONo", obj.ObjPOTMain.str_SONo)
            BaseConn.cmd.Parameters.AddWithValue("@DrawingNo", obj.ObjPOTMain.str_DrawingNo)
            BaseConn.cmd.Parameters.AddWithValue("@Area", obj.ObjPOTMain.str_Area)
            BaseConn.cmd.Parameters.AddWithValue("@Unit", obj.ObjPOTMain.str_Unit)
            BaseConn.cmd.Parameters.AddWithValue("@Qty", obj.ObjPOTMain.str_Qty)
            BaseConn.cmd.Parameters.AddWithValue("@Price", obj.ObjPOTMain.str_Price)
            BaseConn.cmd.Parameters.AddWithValue("@Amount", obj.ObjPOTMain.str_Amount)
            BaseConn.cmd.Parameters.AddWithValue("@POTNo", obj.ObjPOTMain.str_POTNo)
            BaseConn.cmd.Parameters.AddWithValue("@POTDesc", obj.ObjPOTMain.str_POTDesc)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjPOTMain.Str_Flag)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            'BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            'BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.ObjPOTMain.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.ObjPOTMain.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@IssuedBy", obj.ObjPOTMain.str_IssuedBy)
            BaseConn.cmd.Parameters.AddWithValue("@IssuedDate", obj.ObjPOTMain.dtp_IssuedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ReceivedBy", obj.ObjPOTMain.str_ReceivedBy)
            BaseConn.cmd.Parameters.AddWithValue("@FactoryTeam", obj.ObjPOTMain.str_FactoryTeam)

            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objproject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objproject.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objproject.str_ProjectLocation)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", obj.ObjPOTMain.int_StatusCancel)
            BaseConn.cmd.Parameters.AddWithValue("@RevisionNo", obj.ObjPOTMain.int_RevisionNo)

            'BaseConn.cmd.Parameters.AddWithValue("@POT", obj.ObjPOTSub.DT_POT)
            BaseConn.cmd.Parameters.AddWithValue("@POT_Attachment", obj.ObjPOTSub.dt_Attachment)
            BaseConn.cmd.Parameters.AddWithValue("@POT_Section", obj.ObjPOTSub.dt_Section)

            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.cmd.ExecuteNonQuery()
            str_DocumentNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_CID)
            'ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _strDBPwd, obj.int_BusinessPeriodID, obj.ObjSubmittalLogMain.str_CreatedBy, obj.ObjStkAdjMain.dtp_CreatedDate, "", "StockAdjustment", Err.Number, "Error in " & obj.ObjStkAdjCommon.str_Flag & " : " & obj.ObjStkAdjMain.str_DocumentNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_POT = _ErrString
    End Function

    Public Sub Get_POTStatusUpdate(ByRef Obj As csPOT, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[FE_GetPOTStatusUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", Obj.ObjPOTMain.Str_VouNo)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.ObjPOTSub.DT_POT = ds.Tables(0)

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_POTStatusUpdate(ByVal obj As csPOT, ByRef str_DocumentNo As String, ByRef intRevNo As Integer, ByVal _StrDBPath As String, ByVal _strDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[FE_POTStatusUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@DT", obj.ObjPOTSub.dt_Section)
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_CID)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_POTStatusUpdate = _ErrString
    End Function

    Public Function Get_LoadApprovedItems(ByRef _VouNo As String, ByVal _Flag As String, ByVal _SiteID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String) As DataTable
        ErrNo = 0
        ErrStr = ""
        dt = New DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[FE_GetPOTDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", 0)
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
End Class
