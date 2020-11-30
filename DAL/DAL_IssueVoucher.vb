'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes

Public Class DAL_IssueVoucher
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csIssueVoucher, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetIssueVoucherDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.ObjIssueVoucherMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", Obj.ObjIssueVoucherMain.str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.ObjIssueVoucherMain.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            If Obj.ObjIssueVoucherMain.str_Flag <> "JO" Then
                Obj.ObjIssueVoucherMain.str_JONo = ds.Tables(0).Rows(0)("JONo").ToString()
                Obj.ObjIssueVoucherMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.ObjIssueVoucherMain.dtp_VouDate = ds.Tables(0).Rows(0)("VouDate").ToString()
                Obj.ObjIssueVoucherMain.dtp_ReturnDate = ds.Tables(0).Rows(0)("ReturnDate").ToString()

                Obj.ObjIssueVoucherMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.ObjIssueVoucherMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
                Obj.ObjIssueVoucherMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.ObjIssueVoucherMain.bool_Status = ds.Tables(0).Rows(0)("Status").ToString()
                Obj.ObjIssueVoucherMain.str_ReqFormNo = ds.Tables(0).Rows(0)("ReqFormNo").ToString()
                Obj.ObjIssueVoucherMain.str_IssuedBy = ds.Tables(0).Rows(0)("IssuedBy").ToString()
                Obj.ObjIssueVoucherMain.str_ReturnedBy = ds.Tables(0).Rows(0)("ReturnedBy").ToString()
                Obj.ObjIssueVoucherMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()
                Obj.ObjIssueVoucherMain.str_ProductionUnitNo = ds.Tables(0).Rows(0)("ProductionUnitNo").ToString()
                Obj.ObjIssueVoucherMain.str_DstLedger = ds.Tables(0).Rows(0)("DstLedger").ToString()
                Obj.ObjIssueVoucherMain.str_DstLedgerDesc = ds.Tables(0).Rows(0)("DstLedgerDesc").ToString()

                Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
                Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
                Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")
            Else
                Obj.ObjIssueVoucherMain.dtp_VouDate = Date.Now
                Obj.ObjIssueVoucherMain.dtp_ReturnDate = Date.Now
                Obj.ObjIssueVoucherMain.int_RevNo = 0
                Obj.ObjIssueVoucherMain.str_Comment = ds.Tables(0).Rows(0)("JODesc").ToString()
                Obj.ObjIssueVoucherMain.str_ProductionUnitNo = ds.Tables(0).Rows(0)("ProdUnitName").ToString()
            End If

            Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
            Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
            Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            Obj.ObjIssueVoucherMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()


            'If ds.Tables(1).Rows.Count > 0 Then
            Obj.ObjIssueVoucherSub.dt_IssueVoucherSub = ds.Tables(1)
            'End If

            If ds.Tables.Count > 2 Then
                If ds.Tables(2).Rows.Count > 0 Then
                    Obj.objproject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                    Obj.objproject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                    Obj.objproject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
                End If
            End If

            If ds.Tables.Count > 3 Then
                Obj.DTBatch = ds.Tables(3)
            End If

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try

    End Sub


    Public Function Update_IssueVoucher(ByVal obj As csIssueVoucher, ByRef VouNo As String, ByRef intRevNo As Integer, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("IssueVoucherUpdated", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.int_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.ObjIssueVoucherMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", obj.ObjIssueVoucherMain.str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.ObjIssueVoucherMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@JONo", obj.ObjIssueVoucherMain.str_JONo)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.ObjIssueVoucherMain.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.ObjIssueVoucherMain.str_Alias)
            BaseConn.cmd.Parameters.AddWithValue("@VouDate", obj.ObjIssueVoucherMain.dtp_VouDate)
            BaseConn.cmd.Parameters.AddWithValue("@ReturnDate", obj.ObjIssueVoucherMain.dtp_ReturnDate)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.ObjIssueVoucherMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@Status", obj.ObjIssueVoucherMain.bool_Status)
            BaseConn.cmd.Parameters.AddWithValue("@ReqFormNo", obj.ObjIssueVoucherMain.str_ReqFormNo)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.ObjIssueVoucherMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", obj.ObjIssueVoucherMain.int_LanguageCode)

            BaseConn.cmd.Parameters.AddWithValue("@IssuedBy", obj.ObjIssueVoucherMain.str_IssuedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ReturnedBy", obj.ObjIssueVoucherMain.str_ReturnedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedLevel", obj.ApprovedLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedComment", obj.ApprovedComment)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedHigherLevel", obj.ApprovedHigherLevel)

            BaseConn.cmd.Parameters.AddWithValue("@IssueVoucherPrefix", obj.ObjIssueVoucherMain.str_IssueVoucherPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objproject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objproject.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objproject.str_ProjectLocation)

            BaseConn.cmd.Parameters.AddWithValue("@WHID", obj.ObjIssueVoucherMain.str_WHID)
            BaseConn.cmd.Parameters.AddWithValue("@ProductionUnitNo", obj.ObjIssueVoucherMain.str_ProductionUnitNo)
            BaseConn.cmd.Parameters.AddWithValue("@DstLedger", obj.ObjIssueVoucherMain.str_DstLedger)

            BaseConn.cmd.Parameters.AddWithValue("@IssueVoucherItemDetailsDT", obj.ObjIssueVoucherSub.dt_IssueVoucherSub)
            BaseConn.cmd.Parameters.AddWithValue("@ItemBatchDT", obj.DTBatch)

            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjIssueVoucherMain.str_Flag)
            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            VouNo = BaseConn.cmd.Parameters("@VOuNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.int_CID)
            ObjDalGeneral.Elog_Insert(obj.int_CID, _StrDBPath, _StrDBPwd, obj.ObjIssueVoucherMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "IssueVoucher", Err.Number, "Error in " & obj.ObjIssueVoucherMain.str_Flag & " : " & obj.ObjIssueVoucherMain.str_VouNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1 'Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_IssueVoucher = _ErrString
    End Function
End Class
