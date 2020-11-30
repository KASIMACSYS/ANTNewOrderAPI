'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes

Public Class DAL_StockAdjustment
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csStockAdjustment, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetStockAdjustmentDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@DocumentNo", Obj.ObjStkAdjMain.str_DocumentNo)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.ObjStkAdjMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.ObjStkAdjMain.dtp_DocumentDate = ds.Tables(0).Rows(0)("DocDate").ToString()
            Obj.ObjStkAdjMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.ObjStkAdjMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo")
            Obj.ObjStkAdjMain.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.ObjStkAdjMain.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            Obj.ObjStkAdjMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
            Obj.ObjStkAdjMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()
            If ds.Tables(1).Rows.Count > 0 Then
                Obj.ObjStkAdjSub.dt_StkAdjSub = ds.Tables(1)
            End If

            If ds.Tables(2).Rows.Count > 0 Then
                Obj.objproject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                Obj.objproject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                Obj.objproject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
            Else
                Obj.objproject.str_ProjectID = ""
                Obj.objproject.str_ProjectLocation = ""
                Obj.objproject.str_WorkOrderNo = ""
            End If

            Obj.DTBatch = ds.Tables(3)

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub


    Public Function Update_StkAdj(ByVal obj As csStockAdjustment, ByRef str_DocumentNo As String, ByRef intRevNo As Integer, ByVal _StrDBPath As String, ByVal _strDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("StockAdjustmentUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.ObjStkAdjCommon.str_stkAdjPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.ObjStkAdjMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.ObjStkAdjMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@DocumentNo", obj.ObjStkAdjMain.str_DocumentNo)
            BaseConn.cmd.Parameters.AddWithValue("@EntryDate", obj.ObjStkAdjMain.dtp_EntryDate)
            BaseConn.cmd.Parameters.AddWithValue("@DocDate", obj.ObjStkAdjMain.dtp_DocumentDate)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.ObjStkAdjMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.ObjStkAdjMain.int_RevNo)

            BaseConn.cmd.Parameters.AddWithValue("@WHID", obj.ObjStkAdjMain.str_WHID)

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
            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", obj.ObjStkAdjMain.int_LanguageCode)
            BaseConn.cmd.Parameters.AddWithValue("@StkAdjItemDetailsDT", obj.ObjStkAdjSub.dt_StkAdjSub)
            BaseConn.cmd.Parameters.AddWithValue("@ItemBatchDT", obj.DTBatch)

            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjStkAdjCommon.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@stkAdjPrefix", obj.ObjStkAdjCommon.str_stkAdjPrefix)
            BaseConn.cmd.Parameters.Add("@DocumentNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.cmd.ExecuteNonQuery()
            str_DocumentNo = BaseConn.cmd.Parameters("@DocumentNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _strDBPwd, obj.ObjStkAdjMain.int_BusinessPeriodID, obj.ObjStkAdjMain.str_CreatedBy, obj.ObjStkAdjMain.dtp_CreatedDate, "", "StockAdjustment", Err.Number, "Error in " & obj.ObjStkAdjCommon.str_Flag & " : " & obj.ObjStkAdjMain.str_DocumentNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_StkAdj = _ErrString
    End Function

End Class
