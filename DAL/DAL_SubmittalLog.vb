Imports Classes

Public Class DAL_SubmittalLog

    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csSubmittalLog, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetSumittalLogDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", Obj.ObjSubmittalLogMain.Str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.ObjSubmittalLogMain.dtp_VouDate = ds.Tables(0).Rows(0)("VouDate").ToString()
            'Obj.ObjStkAdjMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.ObjSubmittalLogMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo")
            Obj.ObjSubmittalLogMain.str_SalesOrderNo = ds.Tables(0).Rows(0)("SalesOrderNo")
            Obj.ObjSubmittalLogMain.str_ItemCode = ds.Tables(0).Rows(0)("ItemCode")
            Obj.ObjSubmittalLogMain.Int_Slno = ds.Tables(0).Rows(0)("Slno")
            Obj.LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            Obj.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
            Obj.ObjSubmittalLogMain.strGUID = ds.Tables(0).Rows(0)("VouGUID").ToString()
            'Obj.ObjStkAdjMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()

            'If ds.Tables(1).Rows.Count > 0 Then
            Obj.ObjSubmittalLogSub.dt_SubmittalLog = ds.Tables(1)
            'End If

            If ds.Tables(2).Rows.Count > 0 Then
                Obj.objproject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                Obj.objproject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                Obj.objproject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
            Else
                Obj.objproject.str_ProjectID = ""
                Obj.objproject.str_ProjectLocation = ""
                Obj.objproject.str_WorkOrderNo = ""
            End If
            'If ds.Tables(2).Rows.Count > 0 Then
            '    Obj.objproject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
            '    Obj.objproject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
            '    Obj.objproject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
            'Else
            '    Obj.objproject.str_ProjectID = ""
            '    Obj.objproject.str_ProjectLocation = ""
            '    Obj.objproject.str_WorkOrderNo = ""
            'End If

            'Obj.DTBatch = ds.Tables(3)

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub


    Public Function Update_SubLog(ByVal obj As csSubmittalLog, ByRef str_DocumentNo As String, ByRef intRevNo As Integer, ByVal _StrDBPath As String, ByVal _strDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_SubmittalLogUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.ObjSubmittalLogMain.Str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", obj.ObjSubmittalLogMain.Str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@VouDate", obj.ObjSubmittalLogMain.dtp_VouDate)
            'BaseConn.cmd.Parameters.AddWithValue("@DocDate", obj.ObjSubmittalLogMain.dtp_DocumentDate)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.ObjSubmittalLogMain.Str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.ObjSubmittalLogMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@SalesOrderNo", obj.ObjSubmittalLogMain.str_SalesOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", obj.ObjSubmittalLogMain.str_ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@Slno", obj.ObjSubmittalLogMain.Int_Slno)
            BaseConn.cmd.Parameters.AddWithValue("@VouGUID", obj.ObjSubmittalLogMain.strGUID)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objproject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objproject.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objproject.str_ProjectLocation)

            BaseConn.cmd.Parameters.AddWithValue("@SubmittalLogDT", obj.ObjSubmittalLogSub.dt_SubmittalLog)

            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjSubmittalLogMain.Str_Flag)
            'BaseConn.cmd.Parameters.AddWithValue("@stkAdjPrefix", obj.ObjStkAdjCommon.str_stkAdjPrefix)

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
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            'ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _strDBPwd, obj.int_BusinessPeriodID, obj.ObjSubmittalLogMain.str_CreatedBy, obj.ObjStkAdjMain.dtp_CreatedDate, "", "StockAdjustment", Err.Number, "Error in " & obj.ObjStkAdjCommon.str_Flag & " : " & obj.ObjStkAdjMain.str_DocumentNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_SubLog = _ErrString
    End Function
    Public Function Get_SOItemDetails(ByVal SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal _SONumber As String, ByVal _Flag As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String) As DataTable
        dt = New DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetSOItemforSubmittalLog]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@SalOrdNo", _SONumber)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

End Class
