Imports Classes
Public Class DAL_Production
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csProduction, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetProductionDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.ObjProductionMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.ObjProductionMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@ProdNo", Obj.ObjProductionMain.str_ProdNo)
            BaseConn.cmd.Parameters.AddWithValue("@JONo", Obj.ObjProductionMain.str_JOBNo)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.ObjProductionSub.dt_FGItems = ds.Tables(0)


            If Obj.ObjProductionMain.str_Flag <> "ADD" Then
                Obj.ObjProductionMain.str_JOBNo = ds.Tables(1).Rows(0)("JOBNo").ToString()
                Obj.ObjProductionMain.dtp_VouDate = ds.Tables(1).Rows(0)("VouDate")
                Obj.ObjProductionMain.dtp_EstDate = ds.Tables(1).Rows(0)("EstDate")
                Obj.ObjProductionMain.dtp_CompDate = ds.Tables(1).Rows(0)("CompDate")
                Obj.ObjProductionMain.str_Status = ds.Tables(1).Rows(0)("Status")
                Obj.ObjProductionMain.str_Comment = ds.Tables(1).Rows(0)("Comment").ToString()
                Obj.ObjProductionMain.str_ProdunitName = ds.Tables(1).Rows(0)("Production").ToString()
                Obj.ObjProductionMain.dbl_TCAmount = ds.Tables(1).Rows(0)("TCAmount")
                Obj.ObjProductionMain.dbl_TCDisAmount = ds.Tables(1).Rows(0)("TCDisAmount")
                Obj.ObjProductionMain.dbl_TCDiscountAmount = ds.Tables(1).Rows(0)("TCDiscountAmount")
                Obj.ObjProductionMain.dbl_TCNetAmount = ds.Tables(1).Rows(0)("TCNetAmount")
                Obj.ObjProductionMain.dbl_TCVatAmount = ds.Tables(1).Rows(0)("TCVatAmount")
                Obj.ObjProductionMain.str_LpoNo = ds.Tables(1).Rows(0)("LpoNo")

                Obj.str_LastUpdatedBy = ds.Tables(1).Rows(0)("LastUpdatedBy").ToString()
                Obj.dtp_LastUpdatedDate = ds.Tables(1).Rows(0)("LastUpdatedDate").ToString()
            Else
                Obj.ObjProductionMain.dtp_EstDate = ds.Tables(1).Rows(0)("EstEndDate")
                Obj.ObjProductionMain.str_ProdunitName = ds.Tables(1).Rows(0)("ProdUnitName")
                Obj.ObjProductionMain.str_LpoNo = ds.Tables(1).Rows(0)("LpoNo")
            End If

            If ds.Tables.Count >= 3 Then
                Obj.DTBatch = ds.Tables(2)
            End If

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub


    Public Function Update_Production(ByVal _StrDBPath As String, ByVal _strDBPwd As String, ByVal obj As csProduction, ByRef str_DocumentNo As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("UpdateProduction", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.ObjProductionMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.ObjProductionMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjProductionMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.ObjProductionMain.str_Prefix)

            BaseConn.cmd.Parameters.AddWithValue("@ProdNo", obj.ObjProductionMain.str_ProdNo)
            BaseConn.cmd.Parameters.AddWithValue("@JOBNo", obj.ObjProductionMain.str_JOBNo)
            BaseConn.cmd.Parameters.AddWithValue("@VouDate", obj.ObjProductionMain.dtp_VouDate)
            BaseConn.cmd.Parameters.AddWithValue("@EstDate", obj.ObjProductionMain.dtp_EstDate)
            BaseConn.cmd.Parameters.AddWithValue("@CompDate", obj.ObjProductionMain.dtp_CompDate)
            BaseConn.cmd.Parameters.AddWithValue("@Status", obj.ObjProductionMain.str_Status)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.ObjProductionMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@WHID", obj.ObjProductionMain.str_WHID)
            BaseConn.cmd.Parameters.AddWithValue("@Production", obj.ObjProductionMain.str_ProdunitName)
            BaseConn.cmd.Parameters.AddWithValue("@LpoNo", obj.ObjProductionMain.str_LpoNo)

            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.ObjProductionMain.dbl_TCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", obj.ObjProductionMain.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", obj.ObjProductionMain.dbl_TCDiscountAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", obj.ObjProductionMain.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCVatAmount", obj.ObjProductionMain.dbl_TCVatAmount)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@ProductionDT", obj.ObjProductionSub.dt_FGItems)
            BaseConn.cmd.Parameters.AddWithValue("@ItemBatchDT", obj.DTBatch)

            BaseConn.cmd.Parameters.Add("@ProdNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.cmd.ExecuteNonQuery()
            str_DocumentNo = BaseConn.cmd.Parameters("@ProdNoOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _strDBPwd, 0, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "StockAdjustment", Err.Number, "Error in " & obj.ObjProductionMain.str_Flag & " : " & obj.ObjProductionMain.str_ProdNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_Production = _ErrString
    End Function

End Class
