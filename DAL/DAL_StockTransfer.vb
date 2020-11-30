'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes
Public Class DAL_StockTransfer
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Function Get_Structure(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal str_SiteID As String, ByVal Obj As csStockTransfer) As csStockTransfer
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetStockTransfer]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@TransferID", Obj.ObjStkTrnsMain.str_TransferID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.ObjStkTrnsMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.ObjStkTrnsMain.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.ObjStkTrnsMain.dtp_TransferDate = ds.Tables(0).Rows(0)("TransferDate").ToString
            Obj.ObjStkTrnsMain.str_FromWH = ds.Tables(0).Rows(0)("FromWH").ToString()
            Obj.ObjStkTrnsMain.str_ToWH = ds.Tables(0).Rows(0)("ToWH").ToString()
            Obj.ObjStkTrnsMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.ObjStkTrnsMain.str_StockTransferStatus = ds.Tables(0).Rows(0)("StockTransferStatus").ToString()
            Obj.ObjStkTrnsMain.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
            Obj.ObjStkTrnsMain.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
            Obj.ObjStkTrnsMain.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.ObjStkTrnsMain.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            Obj.ObjStkTrnsMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
            Obj.ObjStkTrnsMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            Obj.ObjStkTrnsMain.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
            Obj.ObjStkTrnsMain.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
            Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")
            If ds.Tables(1).Rows.Count > 0 Then
                Obj.ObjStkTrnsSub.dt_StkTrnsSub = ds.Tables(1)
            End If
            Obj.DTBatch = ds.Tables(2)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Get_Structure = Obj
        Return Get_Structure
    End Function
    Public Function Update_StkTrns(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal str_SiteID As String, ByVal obj As csStockTransfer, ByRef str_DocumentNo As String, ByRef outRevNo As Integer, ByRef ErrNo As Integer, ByRef ErrStr As String) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("StockTransferUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.ObjStkTrnsMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@TransferID", obj.ObjStkTrnsMain.str_TransferID)
            BaseConn.cmd.Parameters.AddWithValue("@TransferDate", obj.ObjStkTrnsMain.dtp_TransferDate)
            BaseConn.cmd.Parameters.AddWithValue("@FromWH", obj.ObjStkTrnsMain.str_FromWH)
            BaseConn.cmd.Parameters.AddWithValue("@ToWH", obj.ObjStkTrnsMain.str_ToWH)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.ObjStkTrnsMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@StockTransferStatus", obj.ObjStkTrnsMain.str_StockTransferStatus)

            BaseConn.cmd.Parameters.AddWithValue("@Prefix", obj.ObjStkTrnsMain.str_Prefix)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.ObjStkTrnsMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.ObjStkTrnsMain.str_MenuID)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.ObjStkTrnsMain.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.ObjStkTrnsMain.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.ObjStkTrnsMain.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.ObjStkTrnsMain.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", obj.ObjStkTrnsMain.int_LanguageCode)
            BaseConn.cmd.Parameters.AddWithValue("@StockTransferDT", obj.ObjStkTrnsSub.dt_StkTrnsSub)
            BaseConn.cmd.Parameters.AddWithValue("@ItemBatchDT", obj.DTBatch)

            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjStkTrnsMain.str_Flag)
            BaseConn.cmd.Parameters.Add("@TransferIDOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            str_DocumentNo = BaseConn.cmd.Parameters("@TransferIDOut").Value.ToString
            outRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.ObjStkTrnsMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "StockTransfer", Err.Number, "Error in '" & obj.ObjStkTrnsMain.str_Flag & "'ED '" & obj.ObjStkTrnsMain.str_TransferID & "' ", ex.Message, 5, 3, 1, 0)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_StkTrns = _ErrString
    End Function
End Class
