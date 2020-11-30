Imports Classes

Public Class DAL_TaxFileAdjustment
    Private ObjDalGeneral As DAL_General
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Public Sub Get_Structure(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef Obj As csTaxFileAdjustment, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetTaxFileAdjustment]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@TaxID", Obj.ObjTaxFileAdjustment.int_TaxID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.ObjTaxFileAdjustment.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.ObjTaxFileAdjustment.int_TaxID = ds.Tables(0).Rows(0)("UID")
            Obj.ObjTaxFileAdjustment.str_VouNo = ds.Tables(0).Rows(0)("VouNo").ToString()
            Obj.ObjTaxFileAdjustment.str_TaxFileVouNo = ds.Tables(0).Rows(0)("TaxFileReturnVouNo").ToString()
            Obj.ObjTaxFileAdjustment.str_TaxDesc = ds.Tables(0).Rows(0)("Description").ToString()
            Obj.ObjTaxFileAdjustment.dtp_VouDate = ds.Tables(0).Rows(0)("VouDate").ToString()
            Obj.ObjTaxFileAdjustment.str_TaxLedgerID = ds.Tables(0).Rows(0)("TaxLedgerID").ToString()
            Obj.ObjTaxFileAdjustment.str_DscLedgerID = ds.Tables(0).Rows(0)("DstLedgerID").ToString()
            Obj.ObjTaxFileAdjustment.dbl_AdjustmentAmt = ds.Tables(0).Rows(0)("AdjustmentAmt").ToString()
            Obj.ObjTaxFileAdjustment.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.ObjTaxFileAdjustment.dt_TaxFileAdjustment = ds.Tables(1)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Function GetTaxLedgerID(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _CID As String, ByVal _Flag As String, ByVal dt_TaxLedgerID As DataTable, ByRef ErrNo As Integer) As DataTable
        Dim _ErrString As String = ""
        ErrNo = 0
        dt = New DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetTaxFileAdjustment]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", 0)
            BaseConn.cmd.Parameters.AddWithValue("@TaxID", 0)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    Public Function Update_TaxFileAdjustment(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal obj As csTaxFileAdjustment, ByRef Int_TaxID As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("UpdateTaxFileAdjustment", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjTaxFileAdjustment.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@TaxID", obj.ObjTaxFileAdjustment.int_TaxID)
            BaseConn.cmd.Parameters.AddWithValue("@TaxVouDate", obj.ObjTaxFileAdjustment.dtp_VouDate)
            BaseConn.cmd.Parameters.AddWithValue("@TaxVouNo", obj.ObjTaxFileAdjustment.str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@TaxFileRtnVouNo", obj.ObjTaxFileAdjustment.str_TaxFileVouNo)
            BaseConn.cmd.Parameters.AddWithValue("@TaxLedgerID", obj.ObjTaxFileAdjustment.str_TaxLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@DstLedgerID", obj.ObjTaxFileAdjustment.str_DscLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Description", obj.ObjTaxFileAdjustment.str_TaxDesc)
            BaseConn.cmd.Parameters.AddWithValue("@TaxAdjustmentAmt", obj.ObjTaxFileAdjustment.dbl_AdjustmentAmt)
            BaseConn.cmd.Parameters.AddWithValue("@TaxComment", obj.ObjTaxFileAdjustment.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.ObjTaxFileAdjustment.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.ObjTaxFileAdjustment.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.ObjTaxFileAdjustment.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.ObjTaxFileAdjustment.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.ObjTaxFileAdjustment.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.ObjTaxFileAdjustment.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.ObjTaxFileAdjustment.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.Add("@TaxIDOut", SqlDbType.Int, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            Int_TaxID = BaseConn.cmd.Parameters("@TaxIDOut").Value
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_CID)
            ObjDalGeneral.Elog_Insert(obj.str_CID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, obj.ObjTaxFileAdjustment.str_CreatedBy, obj.ObjTaxFileAdjustment.dtp_CreatedDate, "", "TaxMaster", Err.Number, "Error in " & obj.ObjTaxFileAdjustment.str_Flag & " : " & obj.ObjTaxFileAdjustment.int_TaxID & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_TaxFileAdjustment = _ErrString
    End Function
End Class
